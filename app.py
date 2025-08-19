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

# ---------- парсер ----------
def parse_invoice(pages_text: list[str]) -> pd.DataFrame:
    # нормализуем строки
    lines: list[str] = []
    for t in pages_text:
        for s in t.splitlines():
            s = " ".join(s.split())
            if s:
                lines.append(s)

    # паттерны
    RE_MPN    = re.compile(r"\b(8\d{10})\b")               # 11 цифр, начинается с 8
    RE_ORDER  = re.compile(r"\b(1\d{5})\b")                # 6 цифр, начинается с 1
    RE_MONEY  = re.compile(r"\d{1,3}(?:[ \u00A0]?\d{3})*[.,]\d{1,2}")  # 1 234,56 / 1234.56
    # qty — строго денежный формат рядом с GAB, НЕ рядом с %
    RE_QTY_TOKEN = re.compile(r"(?<!\d)\d{1,4}[.,]\d{2}(?!\d)")

    def to_float(tok: str) -> float:
        return float(tok.replace(" ", "").replace("\u00A0", "").replace(",", "."))

    def to_int_qty(tok: str) -> int:
        return int(round(to_float(tok)))

    def last_money(s: str) -> str | None:
        toks = RE_MONEY.findall(s)
        return toks[-1] if toks else None

    current_order: str | None = None
    rows = []

    for i, line in enumerate(lines):
        # обновляем order (берём ближайший сверху)
        m_ord = RE_ORDER.search(line)
        if m_ord:
            current_order = m_ord.group(1)

        m_mpn = RE_MPN.search(line)
        if not m_mpn:
            continue
        mpn = m_mpn.group(1)

        # -------- QTY рядом с GAB --------
        qty = None
        # пробуем в текущей строке
        def pick_qty_from_string(s: str) -> int | None:
            if not s:
                return None
            s_low = s.lower()
            pos = s_low.find("gab")
            if pos == -1:
                return None
            # окно вокруг GAB
            window_left  = max(0, pos - 30)
            window_right = min(len(s), pos + 30)
            window = s[window_left:window_right]

            # исключаем проценты (например 58%)
            if "%" in window:
                # но проценты могут быть отдельно; мы всё равно фильтруем токены
                pass

            # ищем денежные токены в окне
            cands = list(RE_QTY_TOKEN.finditer(window))
            if not cands:
                return None
            # берём ближайший к GAB: минимальная дистанция от индекса pos
            best_tok = None
            best_dist = None
            for m in cands:
                if "%" in window[max(0, m.start()-2): m.end()+2]:
                    continue
                dist = min(abs((window_left + m.start()) - pos),
                           abs((window_left + m.end())   - pos))
                if best_dist is None or dist < best_dist:
                    best_dist = dist
                    best_tok = m.group(0)
            if best_tok:
                try:
                    return to_int_qty(best_tok)
                except Exception:
                    return None
            return None

        qty = pick_qty_from_string(line)
        if qty is None and i+1 < len(lines):
            qty = pick_qty_from_string(lines[i+1])
        if qty is None and i > 0:
            qty = pick_qty_from_string(lines[i-1])
        if qty is None:
            qty = 0

        # -------- TOTAL (последняя денежная сумма в строке; если путается с qty, берём предыдущую) --------
        total_tok = last_money(line)
        if not total_tok and i + 1 < len(lines):
            total_tok = last_money(lines[i + 1])

        if total_tok:
            try:
                # если total совпал с qty (например 400,00) и в строке есть ещё суммы — попробуем взять предыдущую
                if abs(to_int_qty(total_tok) - qty) == 0:
                    toks = RE_MONEY.findall(line)
                    if len(toks) >= 2 and toks[-1] == total_tok:
                        # предпочитаем предпоследнюю ТОЛЬКО если она не равна qty
                        prev_tok = toks[-2]
                        if abs(to_int_qty(prev_tok) - qty) != 0:
                            total_tok = prev_tok
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

    if not rows:
        return pd.DataFrame(columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])

    df = pd.DataFrame(rows)
    # уникальность по (Order, MPN), порядок
    df = df.drop_duplicates(subset=["Order reference","MPN"], keep="last")
    df = df.sort_values(["Order reference","MPN"]).reset_index(drop=True)
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
    st.info("1) Залей PDF → 2) проверь предпросмотр → 3) «Скачать Excel».")
