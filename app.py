import streamlit as st
import re, io
import pandas as pd
import fitz  # PyMuPDF
from openpyxl import Workbook, load_workbook

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

# ---------- regex ----------
RE_MPN    = re.compile(r"\b(8\d{10})\b")
RE_ORDER  = re.compile(r"(?:#\s*)?(1\d{5})\b")
RE_MONEY  = re.compile(r"\d{1,3}(?:[ \u00A0]?\d{3})*[.,]\d{2}")   # 1 234,56 | 1234.56
RE_DEC    = re.compile(r"^\d{1,4}[.,]\d{2}$")                    # 7,00 | 400,00

def to_float(tok: str) -> float:
    return float(tok.replace(" ", "").replace("\u00A0", "").replace(",", "."))

def to_int(tok: str) -> int:
    return int(round(to_float(tok)))

# ---------- PDF -> линии слов с координатами ----------
def page_lines_with_words(pdf_bytes: bytes):
    """Возвращает список страниц, где страница = список 'линий',
    а линия = список слов: (x0, y0, x1, y1, text) отсортированных по x."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    all_pages = []
    for p in doc:
        words = p.get_text("words")  # x0,y0,x1,y1, word, block, line, word_no
        # сгруппируем по 'почти одинаковому' y (толеранс)
        words.sort(key=lambda w: (round(w[1], 1), w[0]))
        lines = []
        line = []
        last_y = None
        tol = 1.2  # толеранс по y (pt)
        for w in words:
            x0, y0, x1, y1, text = w[:5]
            if last_y is None or abs(y0 - last_y) <= tol:
                line.append((x0, y0, x1, y1, text))
                last_y = y0 if last_y is None else (last_y + y0) / 2
            else:
                if line:
                    line.sort(key=lambda t: t[0])
                    lines.append(line)
                line = [(x0, y0, x1, y1, text)]
                last_y = y0
        if line:
            line.sort(key=lambda t: t[0])
            lines.append(line)
        all_pages.append(lines)
    return all_pages

# ---------- парсер координатный ----------
def parse_invoice(pdf_bytes: bytes) -> pd.DataFrame:
    pages = page_lines_with_words(pdf_bytes)
    rows = []

    for lines in pages:
        # для поиска order удобно иметь «плоский» текст строки
        plain_lines = [" ".join([w[4] for w in ln]) for ln in lines]

        for i, ln in enumerate(lines):
            # якорь: MPN в этой строке
            texts = [w[4] for w in ln]
            joined = " ".join(texts)
            m_mpn = RE_MPN.search(joined)
            if not m_mpn:
                continue

            mpn = m_mpn.group(1)

            # --- Order: ищем наверх до 5 строк, иначе 1 вниз
            order = ""
            for k in range(i, max(-1, i - 5), -1):
                m_o = RE_ORDER.search(plain_lines[k])
                if m_o:
                    order = m_o.group(1)
                    break
            if not order and i + 1 < len(plain_lines):
                m_o = RE_ORDER.search(plain_lines[i + 1])
                if m_o:
                    order = m_o.group(1)

            # --- Qty: ищем токен сразу справа от 'GAB' в той же линии
            qty = None
            def qty_from_line(line_words):
                # ищем индекс токена 'GAB'
                for idx, w in enumerate(line_words):
                    if "GAB" == w[4].upper():
                        # ищем ближайший справа токен вида 7,00 в пределах 80 pt
                        x_gab = w[2]
                        best = None
                        best_dx = None
                        for j in range(idx + 1, min(idx + 8, len(line_words))):
                            t = line_words[j][4]
                            if RE_DEC.match(t):
                                dx = line_words[j][0] - x_gab
                                if 0 <= dx <= 80:  # рядом по горизонтали
                                    if best_dx is None or dx < best_dx:
                                        best_dx = dx
                                        best = t
                        if best:
                            return to_int(best)
                return None

            qty = qty_from_line(ln)
            # если перенесли на следующую строку — пробуем там, но в той же зоне X
            if qty is None and i + 1 < len(lines):
                qty = qty_from_line(lines[i + 1])
            if qty is None:
                qty = 0

            # --- Total (Summa): последний денежный токен в строке (по X)
            def last_money_from_line(line_words):
                money_tokens = [(w[0], w[4]) for w in line_words if RE_MONEY.fullmatch(w[4])]
                if money_tokens:
                    money_tokens.sort(key=lambda t: t[0])  # по x
                    return money_tokens[-1][1]
                return None

            total_tok = last_money_from_line(ln)
            if not total_tok and i + 1 < len(lines):
                total_tok = last_money_from_line(lines[i + 1])
            total = total_tok or "0,00"

            # защита: если total совпал с qty (например '400,00') и в строке ещё была сумма — возьмём предпоследнюю
            if total_tok and qty:
                try:
                    if abs(to_int(total_tok) - qty) == 0:
                        money_tokens = [(w[0], w[4]) for w in ln if RE_MONEY.fullmatch(w[4])]
                        if len(money_tokens) >= 2:
                            money_tokens.sort(key=lambda t: t[0])
                            alt = money_tokens[-2][1]
                            if abs(to_int(alt) - qty) != 0:
                                total = alt
                except Exception:
                    pass

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
    # дедуп по (Order, MPN)
    df = df.drop_duplicates(subset=["Order reference", "MPN"], keep="last")
    # сортировка: по Order, затем по MPN
    df = df.sort_values(["Order reference", "MPN"]).reset_index(drop=True)
    return df

# ---------- UI ----------
pdf_file = st.file_uploader("Загрузить PDF-счёт", type=["pdf"])
tpl_file = st.file_uploader("Шаблон Excel (необязательно)", type=["xlsx"])

if pdf_file:
    pdf_bytes = pdf_file.read()
    df = parse_invoice(pdf_bytes)

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
    st.info("Залей PDF → спарсим по координатам (GAB → Daudz., крайняя справа сумма → Summa, #1xxxxx → Order).")
