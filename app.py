import streamlit as st, re, io, pandas as pd
from PyPDF2 import PdfReader
import fitz  # PyMuPDF (рендер страниц в картинки для OCR)
import pytesseract
from PIL import Image
from openpyxl import load_workbook, Workbook
from utils import load_profiles, cleanse_mpn

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

# --- Безопасная загрузка профилей правил из YAML
def load_rules_safe():
    try:
        profiles = load_profiles("supplier_profiles.yaml")
        return profiles.get("default", {})
    except Exception:
        # Фоллбэк-правила, если YAML поврежден/отсутствует
        return {
            "remove_leading_C_in_mpn": True,
            "materom_mpn_before_dash": True,
            "order_marker_regex": r"(?i)\bOrder[_\s-]*(\d{4,})",
            "mpn_patterns": [
                r"(?i)C?(\d{2}\.\d{5}-\d{4})",
                r"(?i)C?(\d{2}\.\d{5}-\d{3,5})",
                r"(?i)C?(\d{8,})",
            ],
            "qty_patterns": [
                r"(?:(?:QTY|Daudz\.|Qty)\s*[:\-]?\s*)(\d+[\.,]?\d*)",
                r"(?:\s)(\d{1,5})\s*(?:GAB|UNID|KOM)?\b",
            ],
            "total_patterns": [
                r"(\d+[\.,]\d{2})\s*(?:EUR|€)?\s*$",
            ],
        }

rules = load_rules_safe()
order_re = re.compile(rules.get("order_marker_regex", r"(?i)\bOrder[_\s-]*(\d{4,})"))

pdf_file = st.file_uploader("Загрузить счет (PDF)", type=["pdf"])
tpl_file = st.file_uploader("Шаблон Excel (необязательно)", type=["xlsx"])

# --- Гибридный парсер PDF: текст → при нужде OCR
def parse_pdf(pdf_bytes: bytes) -> pd.DataFrame:
    rows, current_order_digits = [], None

    # 1) Пытаемся вытащить текст напрямую (PyPDF2)
    text_pages = []
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        for page in reader.pages:
            text_pages.append(page.extract_text() or "")
    except Exception:
        text_pages = []

    def too_small(pages):
        # если на страницах почти нет текста — вероятно скан
        return not any(len(p.strip()) > 50 for p in pages)

    # 2) Если текста мало — включаем OCR (PyMuPDF + Tesseract)
    if too_small(text_pages):
        text_pages = []
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for page in doc:
            pix = page.get_pixmap(dpi=200)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            # Если установлены языковые пакеты — используем их; иначе упадём на eng
            try:
                ocr_txt = pytesseract.image_to_string(img, lang="eng+rus+lav")
            except Exception:
                ocr_txt = pytesseract.image_to_string(img, lang="eng")
            text_pages.append(ocr_txt)

    # Вспомогательная функция: найти первое совпадение по списку паттернов из YAML
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

    # 3) Разбираем построчно
    for text in text_pages:
        for raw_line in text.splitlines():
            line = " ".join(raw_line.split())  # нормализуем пробелы

            # Смена блока заказа: Order_123456_...
            m_order = order_re.search(line)
            if m_order:
                current_order_digits = m_order.group(1)

            # MPN
            mpn = find_first("mpn_patterns", line)
            if not mpn:
                continue
            mpn = cleanse_mpn(mpn, rules)  # удаляем 'C' и пр. нормы

            # Quantity
            def to_int(x): return int(float(x.replace(",", ".")))
            qty = find_first("qty_patterns", line, to_int)
            if qty is None:
                # fallback: число перед ценой в конце строки
                m = re.search(r"\b(\d{1,4})\b.*(\d+[.,]\d{2})\s*(?:EUR|€)?\s*$", line)
                if m:
                    qty = int(m.group(1))
            if qty is None:
                qty = 1

            # Totalsprice
            def to_money(x): return round(float(x.replace(",", ".")), 2)
            total = find_first("total_patterns", line, to_money)
            if total is None:
                total = 0.0

            rows.append([mpn, "", qty, total, current_order_digits or ""])

    return pd.DataFrame(rows, columns=["MPN", "Replacem", "Quantity", "Totalsprice", "Order reference"])


# --- UI
if pdf_file:
    df = parse_pdf(pdf_file.read())
    st.subheader("Предпросмотр")
    st.dataframe(df, use_container_width=True)

    st.caption("Можно исправлять и добавлять строки ниже:")
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)

    if st.button("Скачать waybill"):
        # берём твой шаблон, либо создаём пустой с заголовками
        if tpl_file:
            wb = load_workbook(tpl_file)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws["A1"] = "MPN"
            ws["B1"] = "Replacem"
            ws["C1"] = "Quantity"
            ws["D1"] = "Totalsprice"
            ws["E1"] = "Order reference"

        # очистка области и запись со 2-й строки
        for r in range(2, 1000):
            for c in range(1, 6):
                ws.cell(row=r, column=c, value=None)

        for i, row in enumerate(edited.values.tolist(), start=2):
            for j, value in enumerate(row, start=1):
                ws.cell(row=i, column=j, value=value)

        bio = io.BytesIO()
        wb.save(bio)
        st.download_button(
            "Скачать waybill.xlsx",
            data=bio.getvalue(),
            file_name="waybill.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Загрузите PDF, чтобы увидеть предпросмотр.")
