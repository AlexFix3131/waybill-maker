import io
import re
from dataclasses import dataclass
from typing import List, Tuple, Optional

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook

# ---------------- UI ----------------
st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

# ---------------- Regex ----------------
RE_MPN      = re.compile(r"\b(8\d{10})\b")                                  # 11 цифр, начинается с 8
RE_MONEY    = re.compile(r"\d{1,3}(?:[ \u00A0]?\d{3})*[.,]\d{2}")           # 1 234,56 | 1234.56
RE_DEC      = re.compile(r"^\d{1,6}[.,]\d{2}$")                             # 7,00 | 400,00
RE_HDR_ART  = re.compile(r"(?i)artik|artikul")                              # Artikuls
RE_HDR_QTY  = re.compile(r"(?i)daudz")                                      # Daudz.
RE_HDR_SUM  = re.compile(r"(?i)summa|summ")                                 # Summa

# «умные» шаблоны заказа (берём чистые 6 цифр начиная с 1):
RE_ORDER_PATTERNS = [
    re.compile(r"(?:^|\s)#\s*(1\d{5})(?:\s|$)"),                            # #125576
    re.compile(r"(?i)order[_\-\s]*0*(1\d{5})"),                             # Order_125867_31.07.25 → 125867
    re.compile(r"(?<![\d.,])(1\d{5})(?![\d.,])"),                           # отдельно стоящее 1xxxxx без пунктуации
]

def to_float(tok: str) -> float:
    return float(tok.replace(" ", "").replace("\u00A0", "").replace(",", "."))

def to_int(tok: str) -> int:
    return int(round(to_float(tok)))

@dataclass
class Word:
    x0: float
    y0: float
    x1: float
    y1: float
    text: str

@dataclass
class ColumnBand:
    name: str
    x_left: float
    x_right: float

# ---------------- PDF helpers ----------------
def load_words_per_page(pdf_bytes: bytes) -> List[List[Word]]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_words: List[List[Word]] = []
    for p in doc:
        words = p.get_text("words")  # x0,y0,x1,y1,text,block,line,word_no
        ws = [Word(w[0], w[1], w[2], w[3], w[4]) for w in words]
        ws.sort(key=lambda w: (round(w.y0, 1), w.x0))
        pages_words.append(ws)
    return pages_words

def group_lines(words: List[Word], y_tol: float = 1.2) -> List[List[Word]]:
    lines: List[List[Word]] = []
    cur: List[Word] = []
    last_y = None
    for w in words:
        if last_y is None or abs(w.y0 - last_y) <= y_tol:
            cur.append(w); last_y = w.y0 if last_y is None else (last_y + w.y0) / 2
        else:
            if cur:
                cur.sort(key=lambda t: t.x0)
                lines.append(cur)
            cur = [w]; last_y = w.y0
    if cur:
        cur.sort(key=lambda t: t.x0)
        lines.append(cur)
    return lines

def find_header_bands(lines: List[List[Word]]) -> Optional[List[ColumnBand]]:
    """
    Ищем строку-шапку (Artikuls / Daudz. / Summa), строим окна колонок по X.
    """
    for ln in lines[:50]:
        line_text = " ".join(w.text for w in ln)
        if RE_HDR_ART.search(line_text) and RE_HDR_QTY.search(line_text) and RE_HDR_SUM.search(line_text):
            # центры меток
            def center(pattern):
                xs = [ (w.x0+w.x1)/2 for w in ln if pattern.search(w.text) ]
                return sum(xs)/len(xs) if xs else None
            cx_art = center(RE_HDR_ART)
            cx_qty = center(RE_HDR_QTY)
            cx_sum = center(RE_HDR_SUM)
            centers = [("Artikuls", cx_art), ("Daudz.", cx_qty), ("Summa", cx_sum)]
            centers = [(n, c) for n, c in centers if c is not None]
            centers.sort(key=lambda t: t[1])
            if len(centers) < 2:  # слабая шапка
                break
            # границы — середины между центрами
            bands: List[ColumnBand] = []
            for i, (name, cx) in enumerate(centers):
                left = (centers[i-1][1] + cx)/2 if i>0 else cx - 70
                right = (cx + centers[i+1][1])/2 if i < len(centers)-1 else cx + 140
                bands.append(ColumnBand(name, left, right))
            # приведём к фиксированным именам по позиции
            bands.sort(key=lambda b: b.x_left)
            for b, nm in zip(bands, ["Artikuls","Daudz.","Summa"]):
                b.name = nm
            return bands
    return None

def words_in_band(line: List[Word], band: ColumnBand) -> List[Word]:
    return [w for w in line if (w.x0 + w.x1)/2 >= band.x_left and (w.x0 + w.x1)/2 <= band.x_right]

# ---------------- Order detection ----------------
def extract_order_from_text(text: str) -> Optional[str]:
    for pat in RE_ORDER_PATTERNS:
        m = pat.search(text)
        if m:
            return m.group(1)
    return None

def find_order_for_line(lines_text: List[str], i: int, lookback: int = 10) -> str:
    """
    Для строки i ищем ПОСЛЕДНЕЕ упоминание заказа в окне [i-lookback, i-1].
    Если не нашли — смотрим строку ниже (i+1).
    """
    start = max(0, i - lookback)
    for j in range(i-1, start-1, -1):
        o = extract_order_from_text(lines_text[j])
        if o:
            return o
    if i + 1 < len(lines_text):
        o = extract_order_from_text(lines_text[i+1])
        if o:
            return o
    return ""

# ---------------- Core extraction ----------------
def parse_pdf_to_df(pdf_bytes: bytes) -> pd.DataFrame:
    pages = load_words_per_page(pdf_bytes)
    out = []

    for page_words in pages:
        lines = group_lines(page_words)
        lines_text = [" ".join(w.text for w in ln) for ln in lines]

        bands = find_header_bands(lines)
        if not bands:
            # без шапки — откажемся, чтобы не плодить ошибки
            continue
        band_map = {b.name: b for b in bands}

        # после шапки начинаем собирать товары
        start_collect = False
        for i, ln in enumerate(lines):
            txt = lines_text[i]

            if not start_collect:
                if (RE_HDR_ART.search(txt) and RE_HDR_QTY.search(txt) and RE_HDR_SUM.search(txt)):
                    start_collect = True
                continue

            # в колонке Artikuls ищем MPN
            mpn = None
            for w in words_in_band(ln, band_map["Artikuls"]):
                m = RE_MPN.search(w.text)
                if m:
                    mpn = m.group(1); break
            if not mpn:
                # fallback: во всей строке (иногда номер клеят левее/правее)
                m = RE_MPN.search(txt)
                if not m:
                    continue
                mpn = m.group(1)

            # qty из Daudz.: первый "7,00/400,00"
            qty = 0
            band_qty_words = words_in_band(ln, band_map["Daudz."])
            if not band_qty_words and i+1 < len(lines):
                band_qty_words = words_in_band(lines[i+1], band_map["Daudz."])
            for w in band_qty_words:
                if RE_DEC.match(w.text):
                    qty = to_int(w.text); break

            # total из Summa: самый правый денежный токен в окне
            totals_words = words_in_band(ln, band_map["Summa"])
            money = [(w.x0, w.text) for w in totals_words if RE_MONEY.fullmatch(w.text)]
            if not money and i+1 < len(lines):
                totals_words2 = words_in_band(lines[i+1], band_map["Summa"])
                money = [(w.x0, w.text) for w in totals_words2 if RE_MONEY.fullmatch(w.text)]
            total_tok = None
            if money:
                money.sort(key=lambda t: t[0])
                total_tok = money[-1][1]
            if not total_tok:
                # крайний правый денежный по всей строке
                money2 = [(w.x0, w.text) for w in ln if RE_MONEY.fullmatch(w.text)]
                if money2:
                    money2.sort(key=lambda t: t[0])
                    total_tok = money2[-1][1]
            total = total_tok or "0,00"

            # если total = "400,00" и совпало с qty (400) — попробуем взять предпредпоследнюю сумму в колонке
            if total_tok and qty:
                try:
                    if abs(to_int(total_tok) - qty) == 0:
                        mm = [(w.x0, w.text) for w in totals_words if RE_MONEY.fullmatch(w.text)]
                        if len(mm) >= 2:
                            mm.sort(key=lambda t: t[0])
                            alt = mm[-2][1]
                            if abs(to_int(alt) - qty) != 0:
                                total = alt
                except Exception:
                    pass

            # order — последний вверх по окну
            order = find_order_for_line(lines_text, i, lookback=10)

            out.append({
                "MPN": mpn,
                "Replacem": "",
                "Quantity": qty,
                "Totalsprice": total,
                "Order reference": order
            })

    if not out:
        return pd.DataFrame(columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])

    df = pd.DataFrame(out)
    df = df.drop_duplicates(subset=["Order reference","MPN"], keep="last")
    df = df.sort_values(["Order reference","MPN"]).reset_index(drop=True)
    return df

# ---------------- UI flow ----------------
pdf_file = st.file_uploader("Загрузить PDF-счёт", type=["pdf"])
tpl_file = st.file_uploader("Шаблон Excel (необязательно)", type=["xlsx"])

if pdf_file:
    pdf_bytes = pdf_file.read()
    df = parse_pdf_to_df(pdf_bytes)

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
    st.info(
        "Мы парсим PDF по координатам: шапка → окна колонок; "
        "MPN — в Artikuls, Qty — токен вида 7,00 в Daudz., Summa — крайняя справа сумма в Summa; "
        "Order — последний #1xxxxx/Order_1xxxxx выше позиции."
    )
