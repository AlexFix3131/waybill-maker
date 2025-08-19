import streamlit as st
import re, io
import pandas as pd
from PyPDF2 import PdfReader
from openpyxl import Workbook, load_workbook

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

# ---------- чтение текста из PDF ----------
def read_pdf_text(pdf_bytes: bytes) -> list[str]:
    pages = []
    reader = PdfReader(io.BytesIO(pdf_bytes))
    for p in reader.pages:
        try:
            t = p.extract_text() or ""
        except Exception:
            t = ""
        pages.append(t)
    return pages

# ---------- парсер под твою таблицу ----------
def parse_invoice(pages_text: list[str]) -> pd.DataFrame:
    # сплющим в строки без пустых
    lines: list[str] = []
    for t in pages_text:
        for s in t.splitlines():
            s = " ".join(s.split())
            if s:
                lines.append(s)

    # паттерны
    RE_MPN   = re.compile(r"\b(8\d{10})\b")             # MPN = 11 цифр, начинается с 8
    RE_ORDER = re.compile(r"\b(1\d{5})\b")              # Order = 6 цифр, начинается с 1
    RE_MONEY = re.compile(r"\d{1,3}(?:[ \u00A0]?\d{3})*[.,]\d{1,2}")  # 1 234,56  |  1234.56
    RE_GAB1  = re.compile(r"(?i)\bGAB\b[^\d%\-]{0,6}(\d+)")
    RE_GAB2  = re.compile(r"(?i)(\d+)\s*\bGAB\b")

    current_order: str | None = None
    rows = []

    def last_money(s: str) -> str | None:
        toks = RE_MONEY.findall(s)
        return toks[-1] if toks else None

    def norm_qty_token(tok: str) -> int:
        return int(float(tok.replace(" ", "").replace(",", ".").replace("\u00A0", "")))

    for i, line in enumerate(lines):
        # обновляем order (берём ближайший сверху)
        m_ord = RE_ORDER.search(line)
        if m_ord:
            current_order = m_ord.group(1)

        # якорь — строка с MPN
        m_mpn = RE_MPN.search(line)
        if not m_mpn:
            continue
        mpn = m_mpn.group(1)

        # ---- qty (рядом с GAB) ----
        qty = None
        for look in (line, lines[i-1] if i > 0 else "", lines[i+1] if i+1 < len(lines) else ""):
            if not look:
                continue
            m = RE_GAB1.search(look) or RE_GAB2.search(look)
            if m:
                try:
                    qty = int(m.group(1))
                    break
                except Exception:
                    pass
        if qty is None:
            qty = 0  # если не нашли GAB — пусть будет 0, чтобы ты это видел

        # ---- total (последняя сумма в строке, если нет — в следующей) ----
        total_tok = last_money(line)
        if not total_tok and i + 1 < len(lines):
            total_tok = last_money(lines[i + 1])

        # защита от путаницы с qty: не берём ...,"400,00" если qty == 400
        if total_tok:
            try:
                if abs(norm_qty_token(total_tok) - qty) == 0 and len(RE_MONEY.findall(line)) > 1:
                    # берём предпоследнюю сумму (обычно "Cena"), последняя будет "Summa"
                    toks = RE_MONEY.findall(line)
                    total_tok = toks[-1] if toks[-1] != f"{qty},00" else toks[-2]
            except Exception:
                pass

        total = total_tok or "0,00"

        rows.append({
            "MPN": mpn,
            "Replacem": "",
            "Quantity": qty,
            "Totalsprice": total,
            "Order reference": current_order or ""
        })

    # уникальность по (Order, MPN) и порядок
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])
    df = df.drop_duplicates(subset=["Order reference", "MPN"], keep="last")
    df = df.sort_values(["Order reference", "MPN"]).reset_index(drop=True)
    return df

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
        st.download_button(
            "Скачать waybill.xlsx", bio.getvalue(),
            file_name="waybill.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("1) Залей PDF\n2) Проверяй предпросмотр\n3) Жми «Скачать Excel»")
