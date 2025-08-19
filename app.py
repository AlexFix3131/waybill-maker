import streamlit as st
import re, io
import pandas as pd
from PyPDF2 import PdfReader
from openpyxl import Workbook, load_workbook

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

# ===== Чтение текста из PDF =====
def get_text(pdf_bytes: bytes) -> list[str]:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    pages = []
    for page in reader.pages:
        try:
            txt = page.extract_text()
            if txt:
                pages.append(txt)
        except:
            continue
    return pages

# ===== Парсер =====
def parse_pdf(text_pages: list[str]) -> pd.DataFrame:
    lines = []
    for t in text_pages:
        for l in t.splitlines():
            l = l.strip()
            if l:
                lines.append(l)

    data = []
    order_ref = None

    for i, line in enumerate(lines):
        # Находим номер заказа (6 цифр, начинается с 1)
        m_order = re.search(r"\b(1\d{5})\b", line)
        if m_order:
            order_ref = m_order.group(1)

        # Находим MPN (11 цифр, начинается с 8)
        m_mpn = re.search(r"\b(8\d{10})\b", line)
        if not m_mpn:
            continue
        mpn = m_mpn.group(1)

        # Quantity (рядом с GAB)
        qty = None
        for look in [line, lines[i-1] if i > 0 else "", lines[i+1] if i+1 < len(lines) else ""]:
            m_qty = re.search(r"\bGAB[^\d]{0,3}(\d+)\b", look, re.I) or re.search(r"\b(\d+)\s*GAB\b", look, re.I)
            if m_qty:
                qty = int(m_qty.group(1))
                break
        if qty is None:
            qty = 0

        # Total price (последнее число с , или .)
        m_price = re.findall(r"\d+[.,]\d{1,2}", line)
        total = m_price[-1] if m_price else "0"

        data.append({
            "MPN": mpn,
            "Replacem": "",
            "Quantity": qty,
            "Totalsprice": total,
            "Order reference": order_ref or ""
        })

    df = pd.DataFrame(data)
    return df.drop_duplicates(subset=["MPN","Order reference"]).reset_index(drop=True)

# ===== Интерфейс =====
pdf_file = st.file_uploader("Загрузить PDF", type=["pdf"])
tpl_file = st.file_uploader("Шаблон Excel (необязательно)", type=["xlsx"])

if pdf_file:
    pages_text = get_text(pdf_file.read())
    df = parse_pdf(pages_text)

    st.subheader("Предпросмотр")
    st.dataframe(df, use_container_width=True)

    if st.button("Скачать Excel"):
        if tpl_file:
            wb = load_workbook(tpl_file)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.append(["MPN","Replacem","Quantity","Totalsprice","Order reference"])

        for r in df.itertuples(index=False):
            ws.append(list(r))

        bio = io.BytesIO()
        wb.save(bio)
        st.download_button("Скачать waybill.xlsx", bio.getvalue(),
                           "waybill.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
