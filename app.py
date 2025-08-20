import io
import re
import statistics
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

# Заказ: и #125576, и Order_125867_..., и просто 125450 без пунктуации
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
    for ln in lines[:60]:
        text = " ".join(w.text for w in ln)
        if RE_HDR_ART.search(text) and RE_HDR_QTY.search(text) and RE_HDR_SUM.search(text):
            def center(pattern):
                xs = [ (w.x0+w.x1)/2 for w in ln if pattern.search(w.text) ]
                return sum(xs)/len(xs) if xs else None
            cx_art = center(RE_HDR_ART)
            cx_qty = center(RE_HDR_QTY)
            cx_sum = center(RE_HDR_SUM)
            centers = [("Artikuls", cx_art), ("Daudz.", cx_qty), ("Summa", cx_sum)]
            centers = [(n, c) for n, c in centers if c is not None]
            centers.sort(key=lambda t: t[1])
            if len(centers) < 2:
                break
            bands: List[ColumnBand] = []
            for i, (name, cx) in enumerate(centers):
                left = (centers[i-1][1] + cx)/2 if i>0 else cx - 80
                right = (cx + centers[i+1][1])/2 if i < len(centers)-1 else cx + 160
                bands.append(ColumnBand(name, left, right))
            bands.sort(key=lambda b: b.x_left)
            for b, nm in zip(bands, ["Artikuls","Daudz.","Summa"]):
                b.name = nm
            return bands
    return None

def fallback_bands(page_words: List[Word]) -> List[ColumnBand]:
    if not page_words:
        return [
            ColumnBand("Artikuls", 0, 200),
            ColumnBand("Daudz.",   200, 400),
            ColumnBand("Summa",    400, 800),
        ]
    x_min = min(w.x0 for w in page_words)
    x_max = max(w.x1 for w in page_words)
    W = x_max - x_min
    a_r = x_min + 0.45 * W
    d_r = x_min + 0.65 * W
    return [
        ColumnBand("Artikuls", x_min - 10, a_r),
        ColumnBand("Daudz.",   a_r,        d_r),
        ColumnBand("Summa",    d_r,        x_max + 20),
    ]

def words_in_band(line: List[Word], band: ColumnBand) -> List[Word]:
    return [w for w in line if (w.x0 + w.x1)/2 >= band.x_left and (w.x0 + w.x1)/2 <= band.x_right]

# ---------------- Order detection ----------------
def extract_order_from_text(text: str) -> Optional[str]:
    for pat in RE_ORDER_PATTERNS:
        m = pat.search(text)
        if m:
            return m.group(1)
    return None

@dataclass
class OrderMarker:
    x: float
    y: float
    value: str

def detect_order_markers(page_words: List[Word]) -> List[OrderMarker]:
    markers: List[OrderMarker] = []
    for w in page_words:
        val = None
        for pat in RE_ORDER_PATTERNS:
            m = pat.search(w.text)
            if m:
                val = m.group(1); break
        if val:
            cx = (w.x0 + w.x1)/2
            cy = (w.y0 + w.y1)/2
            markers.append(OrderMarker(cx, cy, val))
    if not markers:
        return []

    xs = [m.x for m in markers]
    x_med = statistics.median(xs)
    filtered = [m for m in markers if abs(m.x - x_med) <= 35]  # узкая колонка
    if len(filtered) >= max(3, len(markers)//2):
        markers = filtered

    markers.sort(key=lambda m: m.y)
    return markers

def find_order_for_line_via_column(markers: List[OrderMarker], line_y: float) -> Optional[str]:
    if not markers:
        return None
    below = [m for m in markers if m.y <= line_y + 2]
    if below:
        return below[-1].value
    near = sorted(markers, key=lambda m: abs(m.y - line_y))
    if near and abs(near[0].y - line_y) <= 30:
        return near[0].value
    return None

def find_order_for_block_fallback(lines_text: List[str], i_start: int, i_end: int) -> str:
    # вверх
    start = max(0, i_start - 15)
    for j in range(i_start-1, start-1, -1):
        o = extract_order_from_text(lines_text[j])
        if o:
            return o
    # внутри
    for j in range(i_start, i_end+1):
        o = extract_order_from_text(lines_text[j])
        if o:
            return o
    # вниз
    down_end = min(len(lines_text)-1, i_end + 10)
    for j in range(i_end+1, down_end+1):
        o = extract_order_from_text(lines_text[j])
        if o:
            return o
    return ""

# ---------------- Core extraction ----------------
def parse_pdf_to_df(pdf_bytes: bytes) -> pd.DataFrame:
    pages = load_words_per_page(pdf_bytes)
    out_rows = []
    prev_bands: Optional[List[ColumnBand]] = None

    for page_words in pages:
        lines = group_lines(page_words)
        lines_text = [" ".join(w.text for w in ln) for ln in lines]

        bands = find_header_bands(lines) or prev_bands or fallback_bands(page_words)
        prev_bands = bands
        band_map = {b.name: b for b in bands}

        # Колонка заказов
        order_markers = detect_order_markers(page_words)

        # Якоря блоков (MPN)
        mpn_idxs: List[int] = []
        for idx, ln in enumerate(lines):
            ln_art = words_in_band(ln, band_map["Artikuls"])
            m = RE_MPN.search(" ".join(w.text for w in ln_art)) or RE_MPN.search(lines_text[idx])
            if m:
                mpn_idxs.append(idx)

        for k, i_start in enumerate(mpn_idxs):
            i_end = (mpn_idxs[k+1] - 1) if k+1 < len(mpn_idxs) else (len(lines) - 1)

            # --- MPN ---
            m = RE_MPN.search(lines_text[i_start])
            if not m:
                ln_art = words_in_band(lines[i_start], band_map["Artikuls"])
                m = RE_MPN.search(" ".join(w.text for w in ln_art))
            if not m:
                continue
            mpn = m.group(1)

            # --- ORDER ---
            line_y = statistics.fmean([w.y0 for w in lines[i_start]]) if lines[i_start] else 0.0
            order = find_order_for_line_via_column(order_markers, line_y) or \
                    find_order_for_block_fallback(lines_text, i_start, i_end)

            # --- QTY ---
            qty = 0
            found_qty_tok: Optional[str] = None

            for w in words_in_band(lines[i_start], band_map["Daudz."]):
                if RE_DEC.match(w.text):
                    found_qty_tok = w.text; break

            if not found_qty_tok:
                for ii in range(i_start, i_end+1):
                    for w in words_in_band(lines[ii], band_map["Daudz."]):
                        if RE_DEC.match(w.text):
                            found_qty_tok = w.text; break
                    if found_qty_tok: break

            if not found_qty_tok:
                toks = lines_text[i_start].split()
                for p, t in enumerate(toks):
                    if "GAB" in t.upper():
                        for t2 in toks[p+1:p+6]:
                            if RE_DEC.match(t2):
                                found_qty_tok = t2; break
                        if found_qty_tok: break
                if not found_qty_tok:
                    for ii in range(i_start, i_end+1):
                        toks = lines_text[ii].split()
                        for p, t in enumerate(toks):
                            if "GAB" in t.upper():
                                for t2 in toks[p+1:p+6]:
                                    if RE_DEC.match(t2):
                                        found_qty_tok = t2; break
                                if found_qty_tok: break
                        if found_qty_tok: break

            if found_qty_tok:
                try: qty = to_int(found_qty_tok)
                except Exception: qty = 0

            # --- TOTAL (Summa) ---
            money: List[Tuple[float, str]] = []
            for ii in range(i_start, i_end+1):
                for w in words_in_band(lines[ii], band_map["Summa"]):
                    if RE_MONEY.fullmatch(w.text):
                        money.append((w.x0, w.text))
            total_tok = None
            if money:
                money.sort(key=lambda t: t[0])
                total_tok = money[-1][1]
            else:
                money2 = []
                for ii in range(i_start, i_end+1):
                    for w in lines[ii]:
                        if RE_MONEY.fullmatch(w.text):
                            money2.append((w.x0, w.text))
                if money2:
                    money2.sort(key=lambda t: t[0])
                    total_tok = money2[-1][1]

            total_str = fmt_money(total_tok)

            try:
                if total_tok and abs(to_int(total_tok) - qty) == 0 and len(money) >= 2:
                    alt = money[-2][1]
                    if abs(to_int(alt) - qty) != 0:
                        total_str = fmt_money(alt)
            except Exception:
                pass

            out_rows.append({
                "MPN": mpn,
                "Replacem": "",
                "Quantity": qty,
                "Totalsprice": total_str,
                "Order reference": order
            })

    if not out_rows:
        return pd.DataFrame(columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])

    df = pd.DataFrame(out_rows)
    df = df.drop_duplicates(keep="last")
    df = df[["MPN","Replacem","Quantity","Totalsprice","Order reference"]]
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
        "Order № ловится по «вертикальной колонке» (по X) и по fallback‑поиску. "
        "Qty — колонка Daudz., Total — крайняя справа сумма в Summa."
    )
