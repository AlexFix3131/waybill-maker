# app.py
import io
import re
import statistics
from dataclasses import dataclass
from typing import List, Tuple, Optional

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook

# ===== ПАРАМЕТРЫ, можно трогать при тонкой настройке =====
Y_LINE_TOL_FACTOR   = 0.65   # чувствительность склейки слов в строки по Y (от средней высоты)
ORDER_X_TOL         = 42.0   # радиус X для «колонки заказов»
ORDER_NEAR_Y        = 30.0   # если нет маркера «выше», берём ближайший в окне ±30px
QTY_MAX_LOOK_LINES  = 2      # сколько соседних линий сверху/снизу смотреть, если в своей не нашли
SUM_MAX_LOOK_LINES  = 2
# =========================================================

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

# ------------ Регулярки -------------
RE_MPN   = re.compile(r"\b(8\d{10})\b")                               # 11 цифр, начинается с 8
RE_MONEY = re.compile(r"\d{1,3}(?:[ \u00A0]?\d{3})*[.,]\d{2}")        # 1 234,56 | 1234.56
RE_DEC   = re.compile(r"^\d{1,6}[.,]\d{2}$")                          # 7,00 | 400,00

RE_HDR_ART = re.compile(r"(?i)artik|artikul")                         # Artikuls
RE_HDR_QTY = re.compile(r"(?i)daudz")                                 # Daudz.
RE_HDR_SUM = re.compile(r"(?i)summa|summ")                            # Summa

RE_ORDER_PATTERNS = [
    re.compile(r"(?:^|\s)#\s*(1\d{5})(?:\s|$)"),
    re.compile(r"(?i)\border[_\-\s]*0*(1\d{5})"),
    re.compile(r"(?<![\d.,])(1\d{5})(?![\d.,])"),
]

def to_float(tok: str) -> float:
    return float(tok.replace(" ", "").replace("\u00A0", "").replace(",", "."))

def to_int(tok: str) -> int:
    return int(round(to_float(tok)))

def fmt_money(tok: Optional[str]) -> str:
    if not tok:
        return "0,00"
    t = tok.replace("\u00A0", " ").strip()
    if "." in t and "," not in t:
        try:
            return f"{to_float(t):.2f}".replace(".", ",")
        except Exception:
            return t
    return t

@dataclass
class Word:
    x0: float; y0: float; x1: float; y1: float; text: str

@dataclass
class Line:
    y: float
    words: List[Word]
    text: str

@dataclass
class ColumnBand:
    name: str; x_left: float; x_right: float

@dataclass
class OrderMarker:
    x: float; y: float; value: str

# ------------ PDF helpers -------------
def load_page_words(pdf_bytes: bytes) -> List[List[Word]]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    out: List[List[Word]] = []
    for p in doc:
        raw = p.get_text("words")
        ws = [Word(w[0], w[1], w[2], w[3], w[4]) for w in raw]
        ws.sort(key=lambda w: (round(w.y0, 1), w.x0))
        out.append(ws)
    return out

def group_into_lines(words: List[Word]) -> List[Line]:
    if not words:
        return []
    heights = [w.y1 - w.y0 for w in words if (w.y1 - w.y0) > 0.2]
    h_med = statistics.median(heights) if heights else 8.0
    y_tol = max(1.2, h_med * Y_LINE_TOL_FACTOR)

    lines: List[List[Word]] = []
    cur: List[Word] = []
    last_y = None
    for w in words:
        if last_y is None or abs(w.y0 - last_y) <= y_tol:
            cur.append(w); last_y = w.y0 if last_y is None else (last_y + w.y0)/2
        else:
            cur.sort(key=lambda t: t.x0)
            lines.append(cur)
            cur = [w]; last_y = w.y0
    if cur:
        cur.sort(key=lambda t: t.x0)
        lines.append(cur)

    out: List[Line] = []
    for ln in lines:
        y_c = statistics.fmean([w.y0 for w in ln])
        text = " ".join(w.text for w in ln)
        out.append(Line(y=y_c, words=ln, text=text))
    out.sort(key=lambda L: L.y)
    return out

def detect_header_bands(lines: List[Line]) -> Optional[List[ColumnBand]]:
    for L in lines[:80]:
        t = L.text
        if RE_HDR_ART.search(t) and RE_HDR_QTY.search(t) and RE_HDR_SUM.search(t):
            def center(pattern):
                xs = [(w.x0+w.x1)/2 for w in L.words if pattern.search(w.text)]
                return sum(xs)/len(xs) if xs else None
            cx_art = center(RE_HDR_ART); cx_qty = center(RE_HDR_QTY); cx_sum = center(RE_HDR_SUM)
            centers = [(n,c) for n,c in [("Artikuls",cx_art),("Daudz.",cx_qty),("Summa",cx_sum)] if c is not None]
            if len(centers) < 2:
                continue
            centers.sort(key=lambda t: t[1])
            bands: List[ColumnBand] = []
            for i,(n,cx) in enumerate(centers):
                left  = (centers[i-1][1]+cx)/2 if i>0 else cx-90
                right = (cx+centers[i+1][1])/2 if i < len(centers)-1 else cx+180
                bands.append(ColumnBand(n, left, right))
            # принудительно имена
            for b, nm in zip(sorted(bands, key=lambda b:b.x_left), ["Artikuls","Daudz.","Summa"]):
                b.name = nm
            return bands
    return None

def rough_bands(words: List[Word]) -> List[ColumnBand]:
    if not words:
        return [ColumnBand("Artikuls",0,200), ColumnBand("Daudz.",200,400), ColumnBand("Summa",400,800)]
    x_min = min(w.x0 for w in words); x_max = max(w.x1 for w in words)
    W = x_max - x_min
    return [
        ColumnBand("Artikuls", x_min-10, x_min+0.47*W),
        ColumnBand("Daudz.",   x_min+0.47*W, x_min+0.66*W),
        ColumnBand("Summa",    x_min+0.66*W, x_max+20)
    ]

def in_band(word: Word, band: ColumnBand) -> bool:
    cx = (word.x0+word.x1)/2
    return band.x_left <= cx <= band.x_right

# ------------ Order detection -------------
def extract_order_from_text(text: str) -> Optional[str]:
    for pat in RE_ORDER_PATTERNS:
        m = pat.search(text)
        if m:
            return m.group(1)
    return None

def detect_order_markers_from_lines(lines: List[Line]) -> List[OrderMarker]:
    markers: List[OrderMarker] = []
    for L in lines:
        val = extract_order_from_text(L.text)
        if val:
            xs = [(w.x0+w.x1)/2 for w in L.words if any(p.search(w.text) for p in RE_ORDER_PATTERNS)]
            cx = statistics.median(xs) if xs else (sum(w.x0 for w in L.words)/len(L.words))
            markers.append(OrderMarker(cx, L.y, val))
    if not markers:
        return []
    # колонка по медиане X
    x_med = statistics.median([m.x for m in markers])
    col = [m for m in markers if abs(m.x - x_med) <= ORDER_X_TOL]
    if len(col) >= max(3, len(markers)//2):
        markers = col
    markers.sort(key=lambda m: m.y)
    return markers

def build_order_segments(markers: List[OrderMarker], top_y: float, bottom_y: float):
    """Возвращает [(y_start, y_end, value)] покрывающие всю страницу."""
    if not markers:
        return [(top_y-1e9, bottom_y+1e9, "")]
    segs = []
    for i, m in enumerate(markers):
        y_start = markers[i-1].y if i>0 else top_y-1e9
        y_end   = markers[i+1].y if i+1<len(markers) else bottom_y+1e9
        # для более аккуратного деления — середины между метками
        if i>0:
            y_start = (markers[i-1].y + m.y)/2
        if i+1<len(markers):
            y_end = (m.y + markers[i+1].y)/2
        segs.append((y_start, y_end, m.value))
    # края
    first = markers[0]; last = markers[-1]
    segs.insert(0, (top_y-1e9, (first.y + (markers[1].y if len(markers)>1 else first.y+200))/2, first.value))
    segs.append(((last.y + (markers[-2].y if len(markers)>1 else last.y-200))/2, bottom_y+1e9, last.value))
    # слить перекрытия и отсортировать
    segs.sort(key=lambda s:s[0])
    merged = []
    for s in segs:
        if not merged or s[0] > merged[-1][1]:
            merged.append(list(s))
        else:
            merged[-1][1] = max(merged[-1][1], s[1])
            merged[-1][2] = s[2]
    return [(a,b,v) for a,b,v in merged]

def order_for_y(segments, y):
    for a,b,v in segments:
        if a <= y <= b:
            return v
    # fallback: ближайшая по Y метка
    nearest = min(segments, key=lambda s: min(abs(s[0]-y), abs(s[1]-y))) if segments else None
    return nearest[2] if nearest else ""

# ------------ Основная выборка -------------
def parse_pdf_to_df(pdf_bytes: bytes) -> pd.DataFrame:
    pages_words = load_page_words(pdf_bytes)
    rows = []

    for pw in pages_words:
        if not pw: 
            continue
        lines = group_into_lines(pw)
        top_y   = min(w.y0 for w in pw)
        bottom_y= max(w.y1 for w in pw)

        bands = detect_header_bands(lines)
        if not bands:
            bands = rough_bands(pw)
        band_map = {b.name: b for b in bands}

        # колонка заказов -> сегменты
        markers = detect_order_markers_from_lines(lines)
        segments = build_order_segments(markers, top_y, bottom_y)

        # собрать список индексов «якорных» строк (где MPN)
        mpn_idxs: List[int] = []
        for i, L in enumerate(lines):
            if RE_MPN.search(L.text):
                mpn_idxs.append(i)

        for idx, i in enumerate(mpn_idxs):
            # границы блока до следующего MPN
            i_end = mpn_idxs[idx+1]-1 if idx+1<len(mpn_idxs) else len(lines)-1
            L = lines[i]
            m = RE_MPN.search(L.text)
            if not m:
                continue
            mpn = m.group(1)
            y_line = L.y

            # ORDER по сегментам
            order = order_for_y(segments, y_line)

            # QTY — поиск в колонке Daudz. на ближайшей строке
            qty_val = None
            search_range = [i]
            for d in range(1, QTY_MAX_LOOK_LINES+1):
                if i-d >= 0: search_range.append(i-d)
                if i+d < len(lines): search_range.append(i+d)
            best = (1e9, None)
            for j in search_range:
                for w in lines[j].words:
                    if in_band(w, band_map["Daudz."]) and RE_DEC.match(w.text):
                        dy = abs(lines[j].y - y_line)
                        if dy < best[0]:
                            best = (dy, w.text)
            if best[1]:
                try: qty_val = to_int(best[1])
                except Exception: qty_val = 0
            else:
                qty_val = 0

            # TOTAL — сумма из Summa, ближайшая по Y и правее всех в своей строке
            total_tok = None
            best = (1e9, -1e9, None)  # (|dy|, x, tok)
            search_range = [i]
            for d in range(1, SUM_MAX_LOOK_LINES+1):
                if i-d >= 0: search_range.append(i-d)
                if i+d < len(lines): search_range.append(i+d)
            for j in search_range:
                money_words = [w for w in lines[j].words if in_band(w, band_map["Summa"]) and RE_MONEY.fullmatch(w.text)]
                for w in money_words:
                    dy = abs(lines[j].y - y_line)
                    xr = max(w.x0, w.x1)
                    cand = (dy, -xr, w.text)  # правее → xr больше, но мы берём min(), поэтому -xr
                    if cand < best:
                        best = cand
            total_tok = best[2]
            total_str = fmt_money(total_tok)

            # если total == qty по значению (400,00 ↔ 400) и есть в строке ещё деньги — возьми следующий правее
            if total_tok:
                try:
                    if abs(to_int(total_tok) - qty_val) == 0:
                        # попробуй найти вторую справа сумму
                        money_words = [w for w in L.words if in_band(w, band_map["Summa"]) and RE_MONEY.fullmatch(w.text)]
                        money_words.sort(key=lambda w: max(w.x0,w.x1))
                        if len(money_words) >= 2:
                            total_str = fmt_money(money_words[-1].text)
                except Exception:
                    pass

            rows.append({
                "MPN": mpn,
                "Replacem": "",
                "Quantity": qty_val,
                "Totalsprice": total_str,
                "Order reference": order
            })

    if not rows:
        return pd.DataFrame(columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])

    df = pd.DataFrame(rows).drop_duplicates(keep="last")
    df = df[["MPN","Replacem","Quantity","Totalsprice","Order reference"]]
    # не сортирую, чтобы не ломать порядок — если надо: df.sort_values(["Order reference","MPN"], inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df

# ------------ UI -------------
pdf_file = st.file_uploader("Загрузить PDF-счёт", type=["pdf"])
tpl_file = st.file_uploader("Шаблон Excel (необязательно)", type=["xlsx"])

if pdf_file:
    data = pdf_file.read()
    df = parse_pdf_to_df(data)
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
    st.info("Логика: колонка заказов → сегменты по Y; MPN=8***********; Qty из Daudz., Total из Summa (по ближайшему Y).")
