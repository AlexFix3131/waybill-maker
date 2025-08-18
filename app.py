import streamlit as st, re, io, pandas as pd
from PyPDF2 import PdfReader
import fitz  # PyMuPDF для OCR
import pytesseract
from PIL import Image
from openpyxl import load_workbook, Workbook
from utils import load_profiles, cleanse_mpn

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

# ---------- Правила из YAML с фоллбэком ----------
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
                r"(?i)(\d{8,})\s*$",   # длинные цифры в конце строки
                r"(?i)(\d{8,})",       # запасной вариант
            ],
            "qty_patterns": [
                r"(?i)(\d+[\.,]?\d*)\s*GAB\b",
                r"(?i)GAB\s*(\d+[\.,]?\d*)",
                r"(?i)(?:QTY|Daudz\.|Qty)\s*[:\-]?\s*(\d+[\.,]?\d*)",
                r"(\d{1,5})(?:[,\.]00)?(?=\s*\d{6,}\s*$)",  # число прямо перед MPN в конце строки
            ],
            "total_patterns": [
                r"(?i)(\d{1,3}(?:[\s\u00A0]?\d{3})*[.,]\d{2})(?!.*\d{1,3}(?:[\s\u00A0]?\d{3})*[.,]\d{2})",
            ],
        }

rules = load_rules_safe()
order_re = re.compile(rules.get("order_marker_regex", r"(?i)(?:\bOrder[_\s-]*(\d{4,})|#\s*(\d{4,}))"))

# ---------- Вспомогалки ----------
def money_to_float(x: str) -> float:
    x = x.replace(" ", "").replace("\u00A0", "")
    return round(float(x.replace(",", ".")), 2)

def qty_to_int(x: str) -> int:
    return int(float(x.replace(" ", "").replace(",", ".")))

def extract_order(line: str) -> str | None:
    m = order_re.search(line)
    if m:
        return m.group(1) or m.group(2)
    return None

# ---------- Текст по страницам (если мало текста — OCR) ----------
def get_all_text(pdf_bytes: bytes) -> list[str]:
    texts = []
    # откроем PDF один раз для OCR
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    # попробуем PyPDF2
    reader = None
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
    except Exception:
        reader = None

    for i in range(len(doc)):
        page_text = ""
        if reader:
            try:
                page_text = reader.pages[i].extract_text() or ""
            except Exception:
                page_text = ""
        if len(page_text.strip()) <= 50:
            # OCR для этой страницы
            pix = doc.load_page(i).get_pixmap(dpi=220)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            try:
                page_text = pytesseract.image_to_string(img, lang="eng+rus+lav")
            except Exception:
                page_text = pytesseract.image_to_string(img, lang="eng")
        texts.append(page_text)
    return texts

# ---------- Автопарсер построчно ----------
def auto_parse(text_pages: list[str]) -> pd.DataFrame:
    rows, current_order = [], None
    money_any_re = re.compile(r"\d{1,3}(?:[\s\u00A0]?\d{3})*[.,]\d{2}")

    def find_first(key: str, line: str, conv=None):
        for patt in rules.get(key, []):
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

    # склеиваем переносы (окна по 3 строки)
    flat_lines = []
    for t in text_pages:
        for raw in t.splitlines():
            s = " ".join(raw.split())
            if s:
                flat_lines.append(s)

    windows = []
    for i, s in enumerate(flat_lines):
        w1 = s
        w2 = s + " " + flat_lines[i+1] if i + 1 < len(flat_lines) else s
        w3 = w2 + " " + flat_lines[i+2] if i + 2 < len(flat_lines) else w2
        windows += [w1, w2, w3]

    for line in windows:
        if not line:
            continue

        # Order
        ord_num = extract_order(line)
        if ord_num:
            current_order = ord_num

        # MPN
        mpn = find_first("mpn_patterns", line)
        if not mpn:
            m_tail = re.search(r"(\d{8,})\s*$", line)
            if m_tail:
                mpn = m_tail.group(1)
        if not mpn:
            continue
        mpn = cleanse_mpn(mpn, rules)

        # Quantity
        qty = find_first("qty_patterns", line, qty_to_int)
        if qty is None:
            m_pre = re.search(r"(\d{1,5})(?:[,\.]00)?\s*" + re.escape(mpn) + r"\s*$", line)
            if m_pre:
                try:
                    qty = int(m_pre.group(1))
                except Exception:
                    qty = None
        if qty is None:
            qty = 1

        # Total
        total = find_first("total_patterns", line, money_to_float)
        if total is None:
            all_money = money_any_re.findall(line)
            if all_money:
                last = all_money[-1]
                if last not in ("0,00", "0.00"):
                    try:
                        total = money_to_float(last)
                    except Exception:
                        total = None
        if total is None:
            total = 0.0

        rows.append([mpn, "", qty, total, current_order or ""])

    if rows:
        df = pd.DataFrame(rows, columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])
        df = df.drop_duplicates(subset=["MPN","Order reference","Quantity","Totalsprice"], keep="first")
        return df
    return pd.DataFrame(columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])

# ---------- UI ----------
pdf_file = st.file_uploader("Загрузить счёт (PDF)", type=["pdf"])
tpl_file = st.file_uploader("Шаблон Excel (необязательно)", type=["xlsx"])

if not pdf_file:
    st.info("Загрузите PDF-счёт. Я сам распознаю и соберу Excel по правилам.")
else:
    pages_text = get_all_text(pdf_file.read())
    df = auto_parse(pages_text)

    with st.expander("DEBUG (распознанный текст, первые 3000 символов)", expanded=False):
        st.text("\n\n".join(pages_text)[:3000])

    st.subheader("Предпросмотр (авто)")
    st.dataframe(df, use_container_width=True)

    if st.button("Скачать waybill"):
        if tpl_file:
            wb = load_workbook(tpl_file); ws = wb.active
        else:
            wb = Workbook(); ws = wb.active
            ws["A1"] = "MPN"; ws["B1"] = "Replacem"; ws["C1"] = "Quantity"
            ws["D1"] = "Totalsprice"; ws["E1"] = "Order reference"

        for r in range(2, 3000):
            for c in range(1, 6):
                ws.cell(row=r, column=c, value=None)

        for i, row in enumerate(df.values.tolist(), start=2):
            for j, value in enumerate(row, start=1):
                ws.cell(row=i, column=j, value=value)

        bio = io.BytesIO(); wb.save(bio)
        st.download_button(
            "Скачать waybill.xlsx",
            data=bio.getvalue(),
            file_name="waybill.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
