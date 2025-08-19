import streamlit as st
import re, io
import pandas as pd
from PyPDF2 import PdfReader
from openpyxl import Workbook, load_workbook

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

# -------------------- helpers --------------------
RE_MPN    = re.compile(r"\b(8\d{10})\b")
RE_ORDER  = re.compile(r"(?:#\s*)?(1\d{5})")
RE_MONEY  = re.compile(r"\d{1,3}(?:[ \u00A0]?\d{3})*[.,]\d{2}")  # 1 234,56 / 1234.56
RE_DEC    = re.compile(r"^\d{1,4}[.,]\d{2}$")                   # 7,00 / 400,00

def as_float(tok: str) -> float:
    return float(tok.replace(" ", "").replace("\u00A0", "").replace(",", "."))

def as_qty_int(tok: str) -> int:
    return int(round(as_float(tok)))

def read_pdf_text(pdf_bytes: bytes) -> list[str]:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    pages = []
    for p in reader.pages:
        try:
            pages.append(p.extract_text() or "")
        except Exception:
            pages.append("")
    return pages

# -------------------- core parser --------------------
def parse_invoice(pages_text: list[str]) -> pd.DataFrame:
    # нормализуем строки
    lines: list[str] = []
    for t in pages_text:
        for s in t.splitlines():
            s = " ".join(s.split())
            if s:
                lines.append(s)

    rows = []
    # сканируем
    for i, line in enumerate(lines):
        # ---- MPN якорь ----
        m_mpn = RE_MPN.search(line)
        if not m_mpn:
            continue
        mpn = m_mpn.group(1)

        # ---- Order: #1xxxxx поблизости (3 строки вверх, 1 вниз) ----
        order = ""
        for j in range(i, max(-1, i-3), -1):
            m_o = RE_ORDER.search(lines[j])
            if m_o:
                order = m_o.group(1); break
        if not order and i+1 < len(lines):
            m_o = RE_ORDER.search(lines[i+1])
            if m_o:
                order = m_o.group(1)

        # ---- Qty (Daudz.) — токен после GAB ----
        def qty_from(s: str) -> int | None:
            # токенизируем по пробелам, ищем 'GAB' как отдельный токен или в составе
            toks = s.split()
            for k, tok in enumerate(toks):
                if "GAB" in tok.upper():
                    # следующий токен, который выглядит как 7,00 / 400,00
                    for t in toks[k+1:k+5]:  # максимум 4 шага вправо возле GAB
                        if RE_DEC.match(t):
                            return as_qty_int(t)
                    break
            return None

        qty = qty_from(line)
        if qty is None and i+1 < len(lines):
            qty = qty_from(lines[i+1])
        if qty is None:
            qty = 0

        # ---- Totalsprice (Summa) — последний денежный токен строки (или следующей) ----
        def last_money(s: str) -> str | None:
            arr = RE_MONEY.findall(s)
            return arr[-1] if arr else None

        total_tok = last_money(line)
        if not total_tok and i+1 < len(lines):
            total_tok = last_money(lines[i+1])
        total = total_tok or "0,00"

        rows.append({
            "MPN": mpn,
            "Replacem": "",
            "Quantity": qty,
            "Totalsprice": total,
            "Order reference": order
        })

    if not rows:
        return pd.DataFrame(columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])

    df = pd.DataFrame(rows)
    df = df.drop_duplicates(subset=["Order reference","MPN"], keep="last")
    df = df.sort_values(["Order reference","MPN"]).reset_index(drop=True)
    return df

# -------------------- UI --------------------
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
        st.download_button(
            "Скачать waybill.xlsx",
            data=bio.getvalue(),
            file_name="waybill.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Загрузи PDF. Мы возьмём: MPN (8***********), Daudz. после GAB, Summa как последний столбец, Order = #1xxxxx.")
