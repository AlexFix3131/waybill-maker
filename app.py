import streamlit as st, re, io, pandas as pd
from PyPDF2 import PdfReader
import fitz  # PyMuPDF для OCR-рендера
import pytesseract
from PIL import Image
from openpyxl import load_workbook, Workbook
from utils import load_profiles, cleanse_mpn

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

# ===== YAML rules (safe load) =====
def load_rules_safe():
    try:
        profiles = load_profiles("supplier_profiles.yaml")
        return profiles.get("default", {})
    except Exception:
        return {
            "remove_leading_C_in_mpn": True,
            "materom_mpn_before_dash": True,
            "order_marker_regex": r"(?i)(?:\bOrder[_\s-]*(\d{4,})|#\s*(\d{4,}))",
            "mpn_patterns": [
                r"(?i)C?(\d{2}\.\d{5}-\d{4})",
                r"(?i)C?(\d{2}\.\d{5}-\d{3,5})",
                r"(?i)(\d{8,})\s*$",
                r"(?i)(\d{8,})",
            ],
            "qty_patterns": [
                r"(?i)(\d+[\.,]?\d*)\s*GAB\b",
                r"(?i)GAB\s*(\d+[\.,]?\d*)",
                r"(?:(?i)(?:QTY|Daudz\.|Qty)\s*[:\-]?\s*)(\d+[\.,]?\d*)",
                r"(\d{1,5})(?:[,\.]00)?(?=\s*\d{6,}\s*$)",
            ],
            "total_patterns": [
                r"(?i)(\d{1,3}(?:[\s\u00A0]?\d{3})*[.,]\d{2})(?!.*\d{1,3}(?:[\s\u00A0]?\d{3})*[.,]\d{2})",
            ],
        }

rules = load_rules_safe()
order_re = re.compile(rules.get("order_marker_regex", r"(?i)(?:\bOrder[_\s-]*(\d{4,})|#\s*(\d{4,}))"))

# ===== Uploads =====
pdf_file = st.file_uploader("Загрузить счет (PDF)", type=["pdf"])
tpl_file = st.file_uploader("Шаблон Excel (необязательно)", type=["xlsx"])

# ===== Helpers =====
def money_to_float(x: str) -> float:
    x = x.replace(" ", "").replace("\u00A0", "")
    return round(float(x.replace(",", ".")), 2)

def qty_to_int(x: str) -> int:
    return int(float(x.replace(" ", "").replace(",", ".")))

def extract_order(line: str) -> str | None:
    m = order_re.search(line)
    if m:
        return (m.group(1) or m.group(2))
    return None

# ===== Text extraction per-page (with optional forced OCR) =====
def get_page_texts(pdf_bytes: bytes, force_ocr_all: bool) -> list[str]:
    texts = []
    reader = None
    num_pages = 0
    if not force_ocr_all:
        try:
            reader = PdfReader(io.BytesIO(pdf_bytes))
            num_pages = len(reader.pages)
        except Exception:
            reader = None
            num_pages = 0
    else:
        # будем OCRить все страницы
        pass

    def ocr_page(idx: int) -> str:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page = doc.load_page(idx)
        pix = page.get_pixmap(dpi=220)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        try:
            return pytesseract.image_to_string(img, lang="eng+rus+lav")
        except Exception:
            return pytesseract.image_to_string(img, lang="eng")

    if reader and not force_ocr_all:
        for i in range(num_pages):
            # сначала пробуем как текст
            try:
                t = reader.pages[i].extract_text() or ""
            except Exception:
                t = ""
            if len(t.strip()) > 50:
                texts.append(t)
            else:
                texts.append(ocr_page(i))
    else:
        # OCR всех страниц
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for i in range(len(doc)):
            texts.append(ocr_page(i))

    return texts

# ===== Auto parse with YAML rules =====
def auto_parse(text_pages: list[str]) -> pd.DataFrame:
    rows, current_order = [], None
    money_any_re = re.compile(r"\d{1,3}(?:[\s\u00A0]?\d{3})*[.,]\d{2}")

    def find_first(key: str, line: str, conv=None):
        for patt in rules.get(key, []):
            m = re.search(patt, line)
            if m:
                val = m.group(1)
                if conv:
                    try: return conv(val)
                    except Exception: return None
                return val
        return None

    for text in text_pages:
        for raw in text.splitlines():
            line = " ".join(raw.split())
            m_ord = extract_order(line)
            if m_ord:
                current_order = m_ord

            mpn = find_first("mpn_patterns", line)
            if not mpn:
                continue
            mpn = cleanse_mpn(mpn, rules)

            qty = find_first("qty_patterns", line, qty_to_int)
            if qty is None:
                m_pre = re.search(r"(\d{1,5})(?:[,\.]00)?\s*" + re.escape(mpn) + r"\s*$", line)
                if m_pre:
                    try: qty = int(m_pre.group(1))
                    except: qty = None
            if qty is None: qty = 1

            total = find_first("total_patterns", line, money_to_float)
            if total is None:
                all_money = money_any_re.findall(line)
                if all_money:
                    last = all_money[-1]
                    if last not in ("0,00", "0.00"):
                        try: total = money_to_float(last)
                        except: total = None
            if total is None: total = 0.0

            rows.append([mpn, "", qty, total, current_order or ""])

    return pd.DataFrame(rows, columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])

# ===== Manual regex extraction (from UI) =====
def manual_extract(text_pages: list[str], mpn_pat: str, qty_pat: str, price_pat: str, order_pat: str) -> pd.DataFrame:
    rows, current_order = [], None
    mpn_re = re.compile(mpn_pat) if mpn_pat else None
    qty_re = re.compile(qty_pat) if qty_pat else None
    price_re = re.compile(price_pat) if price_pat else None
    order_re_local = re.compile(order_pat) if order_pat else None

    for text in text_pages:
        for raw in text.splitlines():
            line = " ".join(raw.split())

            if order_re_local:
                m = order_re_local.search(line)
                if m:
                    # берём первую непустую группу
                    for g in m.groups():
                        if g:
                            current_order = g
                            break

            if not mpn_re:
                continue
            m_mpn = mpn_re.search(line)
            if not m_mpn:
                continue
            mpn = cleanse_mpn(m_mpn.group(1), rules)

            qty = 1
            if qty_re:
                m_qty = qty_re.search(line)
                if m_qty:
                    try: qty = qty_to_int(m_qty.group(1))
                    except: qty = qty

            total = 0.0
            if price_re:
                m_pr = price_re.findall(line)
                if m_pr:
                    last = m_pr[-1] if isinstance(m_pr[-1], str) else m_pr[-1][0]
                    try: total = money_to_float(last)
                    except: total = 0.0

            rows.append([mpn, "", qty, total, current_order or ""])

    return pd.DataFrame(rows, columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])

# ===== UI flow =====
if not pdf_file:
    st.info("Загрузите PDF, чтобы увидеть предпросмотр.")
else:
    with st.expander("⚙️ Режим распознавания", expanded=False):
        force_ocr_all = st.checkbox("Force OCR for all pages (распознать все страницы через Tesseract)", value=False)

    # 1) Получаем текст
    pages = get_page_texts(pdf_file.read(), force_ocr_all)
    st.text_area("DEBUG: что удалось вытащить из PDF/OCR (первые 5000 символов)",
                 "\n\n".join(pages)[:5000], height=220)

    # 2) Авторазбор
    df_auto = auto_parse(pages)

    st.subheader("Предпросмотр")
    st.dataframe(df_auto, use_container_width=True)

    # 3) Ручные инструменты
    with st.expander("🔧 Ручные правила (Regex) — вытащить из всего текста", expanded=False):
        colA, colB = st.columns(2)
        mpn_pat = colA.text_input("MPN regex (1-я группа — код)", r"(?i)(\d{8,})\s*$")
        qty_pat = colA.text_input("Quantity regex (1-я группа — число)", r"(?i)(\d+[\.,]?\d*)\s*GAB\b")
        price_pat = colA.text_input("Price regex (все совпадения, берём последнее)", r"(?i)\d{1,3}(?:[\s\u00A0]?\d{3})*[.,]\d{2}")
        order_pat = colB.text_input("Order regex (1-2 группа — номер)", r"(?i)(?:\bOrder[_\s-]*(\d{4,})|#\s*(\d{4,}))")

        if st.button("Extract (manual regex)"):
            df_manual = manual_extract(pages, mpn_pat, qty_pat, price_pat, order_pat)
            st.session_state["df_manual"] = df_manual

        if "df_manual" in st.session_state:
            st.write("Результат manual-извлечения:")
            st.dataframe(st.session_state["df_manual"], use_container_width=True)

    with st.expander("🧾 Отметить строки из полного текста и извлечь", expanded=False):
        # показываем все строки с чекбоксом
        lines = []
        for t in pages:
            lines.extend([l for l in t.splitlines() if l.strip()])
        checked = []
        for i, line in enumerate(lines[:1000]):  # защитимся от очень длинных документов
            if st.checkbox(line, key=f"pick_{i}"):
                checked.append(line)
        if st.button("Извлечь отмеченные (эвристика)"):
            tmp_df = []
            for line in checked:
                line_n = " ".join(line.split())
                ord_id = extract_order(line_n) or ""
                # MPN — длинные цифры в конце
                m_mpn = re.search(r"(\d{8,})\s*$", line_n) or re.search(r"(\d{8,})", line_n)
                if not m_mpn: continue
                mpn = cleanse_mpn(m_mpn.group(1), rules)
                # Qty
                m_qty = re.search(r"(?i)(\d+[\.,]?\d*)\s*GAB\b", line_n) or re.search(r"(?i)GAB\s*(\d+[\.,]?\d*)", line_n)
                qty = qty_to_int(m_qty.group(1)) if m_qty else 1
                # Price — последняя денежка
                m_all = re.findall(r"\d{1,3}(?:[\s\u00A0]?\d{3})*[.,]\d{2}", line_n)
                total = money_to_float(m_all[-1]) if m_all else 0.0
                tmp_df.append([mpn, "", qty, total, ord_id])
            st.session_state["df_marked"] = pd.DataFrame(tmp_df, columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])
        if "df_marked" in st.session_state:
            st.write("Результат из отмеченных строк:")
            st.dataframe(st.session_state["df_marked"], use_container_width=True)

    # 4) Сводная таблица к выгрузке
    out_df = df_auto.copy()
    if "df_manual" in st.session_state and not st.session_state["df_manual"].empty:
        out_df = pd.concat([out_df, st.session_state["df_manual"]], ignore_index=True)
    if "df_marked" in st.session_state and not st.session_state["df_marked"].empty:
        out_df = pd.concat([out_df, st.session_state["df_marked"]], ignore_index=True)

    out_df = out_df.dropna(how="all")
    st.caption("Можно исправлять и добавлять строки ниже:")
    edited = st.data_editor(out_df if not out_df.empty else pd.DataFrame(columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"]),
                            num_rows="dynamic", use_container_width=True)

    # 5) Export
    if st.button("Скачать waybill"):
        if tpl_file:
            wb = load_workbook(tpl_file); ws = wb.active
        else:
            wb = Workbook(); ws = wb.active
            ws["A1"] = "MPN"; ws["B1"] = "Replacem"; ws["C1"] = "Quantity"
            ws["D1"] = "Totalsprice"; ws["E1"] = "Order reference"

        # очистка и запись со 2-й строки
        for r in range(2, 3000):
            for c in range(1, 6):
                ws.cell(row=r, column=c, value=None)

        # пишем текущие данные редактора
        for i, row in enumerate(edited.values.tolist(), start=2):
            for j, value in enumerate(row, start=1):
                ws.cell(row=i, column=j, value=value)

        bio = io.BytesIO(); wb.save(bio)
        st.download_button(
            "Скачать waybill.xlsx",
            data=bio.getvalue(),
            file_name="waybill.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
