import streamlit as st, re, io, pandas as pd
from PyPDF2 import PdfReader
import fitz  # PyMuPDF для OCR-рендера
import pytesseract
from PIL import Image
from openpyxl import load_workbook, Workbook
from utils import load_profiles, cleanse_mpn

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
# отключаем файловый watcher (устраняет inotify warning)
st.set_option('server.fileWatcherType', 'none')

st.title("📦 Waybill Maker")

# --- Безопасная загрузка YAML-правил
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
order_re = re.compile(
    rules.get("order_marker_regex", r"(?i)(?:\bOrder[_\s-]*(\d{4,})|#\s*(\d{4,}))")
)

pdf_file = st.file_uploader("Загрузить счет (PDF)", type=["pdf"])
tpl_file = st.file_uploader("Шаблон Excel (необязательно)", type=["xlsx"])

# --- Гибридный парсер: по каждой странице — текст, при необходимости OCR
def parse_pdf(pdf_bytes: bytes) -> pd.DataFrame:
    rows, current_order_digits = [], None

    # пробуем открыть как текстовый PDF
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        num_pages = len(reader.pages)
    except Exception:
        reader = None
        num_pages = 0

    page_texts = []

    def page_to_text(page_index: int) -> str:
        # 1) текст напрямую
        if reader:
            try:
                t = reader.pages[page_index].extract_text() or ""
                if len(t.strip()) > 50:
                    return t
            except Exception:
                pass
        # 2) OCR для этой страницы
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page = doc.load_page(page_index)
        pix = page.get_pixmap(dpi=220)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        try:
            return pytesseract.image_to_string(img, lang="eng+rus+lav")
        except Exception:
            return pytesseract.image_to_string(img, lang="eng")

    if num_pages > 0:
        for i in range(num_pages):
            page_texts.append(page_to_text(i))
    else:
        # fallback: OCR всего файла
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for p in doc:
            pix = p.get_pixmap(dpi=220)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            try:
                page_texts.append(pytesseract.image_to_string(img, lang="eng+rus+lav"))
            except Exception:
                page_texts.append(pytesseract.image_to_string(img, lang="eng"))

    # --- DEBUG: что распознали (обрезано для интерфейса)
    st.text_area(
        "DEBUG: что удалось вытащить из PDF/OCR",
        "\n\n--- PAGE ---\n\n".join([t[:2000] for t in page_texts]),
        height=260,
    )

    def find_first(pattern_key: str, line: str, conv=None):
        for patt in rules.get(pattern_key, []):
            m = re.search(patt, line)
            if m:
                val = m.group(1)
                if conv:
                    try:
                        return conv(val)
                    except Exception:
                        return None
                return val
        return None

    money_any_re = re.compile(r"\d{1,3}(?:[\s\u00A0]?\d{3})*[.,]\d{2}")

    # разбор построчно
    for text in page_texts:
        for raw_line in text.splitlines():
            line = " ".join(raw_line.split())

            # Order: "#123456" ИЛИ "Order_123456"
            m_order = order_re.search(line)
            if m_order:
                current_order_digits = (m_order.group(1) or m_order.group(2))

            # MPN
            mpn = find_first("mpn_patterns", line)
            if not mpn:
                continue
            mpn = cleanse_mpn(mpn, rules)

            # Quantity
            def to_int(x): 
                return int(float(x.replace(" ", "").replace(",", ".")))
            qty = find_first("qty_patterns", line, to_int)
            if qty is None:
                # число прямо перед MPN в конце
                m_pre = re.search(r"(\d{1,5})(?:[,\.]00)?\s*" + re.escape(mpn) + r"\s*$", line)
                if m_pre:
                    try:
                        qty = int(m_pre.group(1))
                    except Exception:
                        qty = None
            if qty is None:
                qty = 1

            # Totalsprice: берём последнюю ненулевую сумму
            def to_money(x):
                x = x.replace(" ", "").replace("\u00A0", "")
                return round(float(x.replace(",", ".")), 2)

            total = find_first("total_patterns", line, to_money)
            if total is None:
                all_money = money_any_re.findall(line)
                if all_money:
                    last = all_money[-1]
                    if last not in ("0,00", "0.00"):
                        try:
                            total = to_money(last)
                        except Exception:
                            total = None
            if total is None:
                total = 0.0

            rows.append([mpn, "", qty, total, current_order_digits or ""])

    return pd.DataFrame(
        rows,
        columns=["MPN", "Replacem", "Quantity", "Totalsprice", "Order reference"],
    )

# --- UI
if pdf_file:
    df = parse_pdf(pdf_file.read())
    st.subheader("Предпросмотр")
    st.dataframe(df, use_container_width=True)

    st.caption("Можно исправлять и добавлять строки ниже:")
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)

    if st.button("Скачать waybill"):
        if tpl_file:
            wb = load_workbook(tpl_file); ws = wb.active
        else:
            wb = Workbook(); ws = wb.active
            ws["A1"] = "MPN"; ws["B1"] = "Replacem"; ws["C1"] = "Quantity"
            ws["D1"] = "Totalsprice"; ws["E1"] = "Order reference"

        # очистка и запись со 2-й строки
        for r in range(2, 1000):
            for c in range(1, 6):
                ws.cell(row=r, column=c, value=None)

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
else:
    st.info("Загрузите PDF, чтобы увидеть предпросмотр.")
