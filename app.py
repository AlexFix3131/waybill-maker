import streamlit as st, re, io, pandas as pd
from PyPDF2 import PdfReader
from openpyxl import load_workbook, Workbook
from utils import load_profiles, cleanse_mpn

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

profiles = load_profiles("supplier_profiles.yaml")
rules = profiles.get("default", {})

pdf_file = st.file_uploader("Загрузить счет (PDF)", type=["pdf"])
tpl_file = st.file_uploader("Шаблон Excel (необязательно)", type=["xlsx"])

def parse_pdf(pdf_bytes):
    rows, current_order_digits = [], None
    reader = PdfReader(io.BytesIO(pdf_bytes))
    for page in reader.pages:
        text = page.extract_text() or ""
        for raw_line in text.splitlines():
            line = " ".join(raw_line.split())

            # 1) ловим смену блока Order
            m_order = re.search(r"(?i)\bOrder[_\s-]*(\d{4,})", line)
            if m_order:
                current_order_digits = m_order.group(1)

            # 2) ищем MPN по паттернам
            mpn = None
            for patt in rules.get("mpn_patterns", []):
                m = re.search(patt, line)
                if m:
                    mpn = m.group(1)
                    break
            if not mpn:
                continue

            # 3) количество
            qty = None
            for patt in rules.get("qty_patterns", []):
                m = re.search(patt, line)
                if m:
                    try:
                        qty = int(float(m.group(1).replace(",", ".")))
                    except:
                        pass
                    if qty is not None:
                        break
            if qty is None:
                # fallback: число перед ценой
                m = re.search(r"\b(\d{1,4})\b.*(\d+[.,]\d{2})\s*(?:EUR|€)?\s*$", line)
                if m:
                    qty = int(m.group(1))

            # 4) итог по строке
            total = None
            for patt in rules.get("total_patterns", []):
                m = re.search(patt, line)
                if m:
                    try:
                        total = round(float(m.group(1).replace(",", ".")), 2)
                    except:
                        pass
                    if total is not None:
                        break

            mpn = cleanse_mpn(mpn, rules)
            rows.append([mpn, "", qty if qty else 1, total if total else 0.0, current_order_digits or ""])

    return pd.DataFrame(rows, columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])

if pdf_file:
    df = parse_pdf(pdf_file.read())
    st.subheader("Предпросмотр")
    st.dataframe(df, use_container_width=True)
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)

    if st.button("Скачать waybill"):
        if tpl_file:
            wb = load_workbook(tpl_file); ws = wb.active
        else:
            wb = Workbook(); ws = wb.active
            ws["A1"]="MPN"; ws["B1"]="Replacem"; ws["C1"]="Quantity"; ws["D1"]="Totalsprice"; ws["E1"]="Order reference"

        # очистка и запись со 2-й строки
        for r in range(2, 1000):
            for c in range(1,6):
                ws.cell(row=r, column=c, value=None)

        for i, row in enumerate(edited.values.tolist(), start=2):
            for j, value in enumerate(row, start=1):
                ws.cell(row=i, column=j, value=value)

        bio = io.BytesIO(); wb.save(bio)
        st.download_button("Скачать waybill.xlsx", data=bio.getvalue(),
                           file_name="waybill.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Загрузите PDF, чтобы увидеть предпросмотр.")
