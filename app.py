import streamlit as st
import re, io
import pandas as pd
from PyPDF2 import PdfReader
from openpyxl import Workbook, load_workbook

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

# ---------- чтение PDF ----------
def read_pdf_text(pdf_bytes: bytes) -> list[str]:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    return [(p.extract_text() or "") for p in reader.pages]

# ---------- парсер ----------
def parse_invoice(pages_text: list[str]) -> pd.DataFrame:
    lines = []
    for t in pages_text:
        for s in t.splitlines():
            s = " ".join(s.split())
            if s:
                lines.append(s)

    RE_MPN   = re.compile(r"\b(8\d{10})\b")       # MPN = 11 цифр на 8
    RE_ORDER = re.compile(r"#(1\d{5})")           # #123456
    RE_MONEY = re.compile(r"\d{1,3}(?:[ \u00A0]?\d{3})*[.,]\d{1,2}")

    def to_float(tok): return float(tok.replace(" ", "").replace(",", "."))
    def to_int(tok): return int(round(to_float(tok)))

    rows = []
    current_order = None

    for i, line in enumerate(lines):
        # обновляем order
        m_ord = RE_ORDER.search(line)
        if m_ord:
            current_order = m_ord.group(1)

        m_mpn = RE_MPN.search(line)
        if not m_mpn:
            continue
        mpn = m_mpn.group(1)

        # ---- Qty: строго после GAB ----
        qty = 0
        gab_pos = line.find("GAB")
        if gab_pos != -1:
            after = line[gab_pos+3:].strip()
            m_qty = re.match(r"^(\d+[.,]\d{2})", after)
            if m_qty:
                qty = to_int(m_qty.group(1))
        # если в этой строке нет — проверим следующую
        if qty == 0 and i+1 < len(lines):
            nxt = lines[i+1]
            gab_pos = nxt.find("GAB")
            if gab_pos != -1:
                after = nxt[gab_pos+3:].strip()
                m_qty = re.match(r"^(\d+[.,]\d{2})", after)
                if m_qty:
                    qty = to_int(m_qty.group(1))

        # ---- Total: последняя сумма в строке ----
        total = None
        toks = RE_MONEY.findall(line)
        if toks:
            total = toks[-1]
        if not total and i+1 < len(lines):
            toks = RE_MONEY.findall(lines[i+1])
            if toks:
                total = toks[-1]
        total = total or "0,00"

        rows.append({
            "MPN": mpn,
            "Replacem": "",
            "Quantity": qty,
            "Totalsprice": total,
            "Order reference": current_order or ""
        })

    df = pd.DataFrame(rows)
    df = df.drop_duplicates(subset=["Order reference","MPN"], keep="last")
    return df.reset_index(drop=True)

# ---------- UI ----------
pdf_file = st.file_uploader("Загрузить PDF-счёт", type=["pdf"])
tpl_file = st.file_uploader("Шаблон Excel (необязательно)", type=["xlsx"])

if pdf_file:
    pages = read_pdf_text(pdf_file.read())
    df = parse_invoice(pages)

    st.subheader("Предпросмотр")
    st.dataframe(df, use_container_width=True)

    if st.button("Скачать Excel"):
        if tpl_file:
            wb = load_workbook(tpl_file); ws = wb.active
        else:
            wb = Workbook(); ws = wb.active
            ws.append(["MPN","Replacem","Quantity","Totalsprice","Order reference"])
        for r in df.itertuples(index=False):
            ws.append(list(r))
        bio = io.BytesIO(); wb.save(bio)
        st.download_button("Скачать waybill.xlsx", bio.getvalue(),
            file_name="waybill.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("1) Залей PDF → 2) проверь предпросмотр → 3) «Скачать Excel».")
