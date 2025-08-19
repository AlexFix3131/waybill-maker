import io
import re
from dataclasses import dataclass
from typing import List, Tuple, Optional

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

# ===== RegEx =====
RE_MPN     = re.compile(r"\b(8\d{10})\b")
RE_ORDER   = re.compile(r"(?:#\s*)?(1\d{5})\b")
RE_MONEY   = re.compile(r"\d{1,3}(?:[ \u00A0]?\d{3})*[.,]\d{2}")     # 1 234,56 | 1234.56
RE_DEC     = re.compile(r"^\d{1,5}[.,]\d{2}$")                       # 7,00 | 400,00
RE_HEADER1 = re.compile(r"(?i)artik|artikul")                        # Artikuls (латыш)
RE_HEADER2 = re.compile(r"(?i)daudz")                                # Daudz.
RE_HEADER3 = re.compile(r"(?i)summa|summ")                           # Summa

def f_to_float(tok: str) -> float:
    return float(tok.replace(" ", "").replace("\u00A0", "").replace(",", "."))

def f_to_int(tok: str) -> int:
    return int(round(f_to_float(tok)))

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

# ---------- PDF helpers ----------
def load_words_per_page(pdf_bytes: bytes) -> List[List[Word]]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_words: List[List[Word]] = []
    for p in doc:
        words = p.get_text("words")  # (x0,y0,x1,y1, text, block, line, word_no)
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
    Ищем строку‑шапку, где есть Artikuls / Daudz. / Summa.
    Строим 3 окна‑колонки, делим по серединам между центрами слов.
    """
    for ln in lines[:40]:  # в верхней части страницы
        texts = " ".join(w.text for w in ln)
        has_art = RE_HEADER1.search(texts) is not None
        has_dau = RE_HEADER2.search(texts) is not None
        has_sum = RE_HEADER3.search(texts) is not None
        if has_art and has_dau and has_sum:
            # возьмём центры слов-меток
            def center_of(pattern):
                cand = [((w.x0 + w.x1) / 2.0) for w in ln if pattern.search(w.text)]
                return sum(cand) / len(cand) if cand else None

            cx_art = center_of(RE_HEADER1)
            cx_dau = center_of(RE_HEADER2)
            cx_sum = center_of(RE_HEADER3)
            centers = [("Artikuls", cx_art), ("Daudz.", cx_dau), ("Summa", cx_sum)]
            centers = [(n, c) for n, c in centers if c is not None]
            if len(centers) < 2:
                continue
            centers.sort(key=lambda t: t[1])
            # границы — середины между соседями
            bands: List[ColumnBand] = []
            for i, (name, cx) in enumerate(centers):
                if i == 0:
                    left = cx - 60  # чуть шире слева
                else:
                    left = (centers[i - 1][1] + cx) / 2
                if i == len(centers) - 1:
                    right = cx + 120  # правую расширим (Summa)
                else:
                    right = (cx + centers[i + 1][1]) / 2
                bands.append(ColumnBand(name, left, right))
            # Убедимся, что есть все 3 (если нет, дополним разумно)
            names = {b.name for b in bands}
            if "Artikuls" not in names or "Daudz." not in names or "Summa" not in names:
                # попытаемся «назвать» по позициям: слева → Artikuls, середина → Daudz., справа → Summa
                bands.sort(key=lambda b: b.x_left)
                alias = ["Artikuls", "Daudz.", "Summa"]
                for b, nm in zip(bands[:3], alias):
                    b.name = nm
            return bands
    return None

def words_in_band(line: List[Word], band: ColumnBand) -> List[Word]:
    return [w for w in line if (w.x0 + w.x1) / 2.0 >= band.x_left and (w.x0 + w.x1) / 2.0 <= band.x_right]

# ---------- Core extraction ----------
def parse_pdf_to_df(pdf_bytes: bytes) -> pd.DataFrame:
    pages = load_words_per_page(pdf_bytes)
    out_rows = []

    for page_words in pages:
        lines = group_lines(page_words)
        bands = find_header_bands(lines)
        current_order = ""  # «текущий» заказ сверху

        # если шапку не нашли — всё равно попробуем простую эвристику
        if not bands:
            # fallback: разделим на 3 равные полосы по ширине страницы
            if not page_words:
                continue
            x_min = min(w.x0 for w in page_words)
            x_max = max(w.x1 for w in page_words)
            w = (x_max - x_min) / 3
            bands = [
                ColumnBand("Artikuls", x_min - 10, x_min + w),
                ColumnBand("Daudz.",  x_min + w, x_min + 2*w),
                ColumnBand("Summa",   x_min + 2*w, x_max + 20),
            ]

        # идём по строкам после шапки
        start_collect = False
        for ln in lines:
            txt_line = " ".join(w.text for w in ln)
            # обновление order если встречается
            mo = RE_ORDER.search(txt_line)
            if mo:
                current_order = mo.group(1)

            # включаем сбор после строки шапки
            if not start_collect:
                if (RE_HEADER1.search(txt_line) and RE_HEADER2.search(txt_line) and RE_HEADER3.search(txt_line)):
                    start_collect = True
                continue

            # из полос берём данные
            band_map = {b.name: words_in_band(ln, b) for b in bands}
            # MPN — ищем 11‑значный на 8 в колонке Artikuls
            mpn = None
            for w in band_map.get("Artikuls", []):
                m = RE_MPN.search(w.text)
                if m:
                    mpn = m.group(1); break
            if not mpn:
                # fallback: во всей строке
                m = RE_MPN.search(txt_line)
                if m:
                    mpn = m.group(1)
            if not mpn:
                continue  # строка без артикула нам не интересна

            # Qty — токен формата 7,00/400,00 в колонке Daudz.
            qty = 0
            for w in band_map.get("Daudz.", []):
                if RE_DEC.match(w.text):
                    qty = f_to_int(w.text)
                    break

            # Total — самый правый денежный в колонке Summa
            total_tok = None
            sums = [(w.x0, w.text) for w in band_map.get("Summa", []) if RE_MONEY.fullmatch(w.text)]
            if sums:
                sums.sort(key=lambda t: t[0])
                total_tok = sums[-1][1]
            # если пусто — попробуем правые токены во всей строке
            if not total_tok:
                sums2 = [(w.x0, w.text) for w in ln if RE_MONEY.fullmatch(w.text)]
                if sums2:
                    sums2.sort(key=lambda t: t[0]); total_tok = sums2[-1][1]
            total = total_tok or "0,00"

            out_rows.append({
                "MPN": mpn,
                "Replacem": "",
                "Quantity": qty,
                "Totalsprice": total,
                "Order reference": current_order
            })

    if not out_rows:
        return pd.DataFrame(columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])

    df = pd.DataFrame(out_rows)
    # чистим дубли по (Order, MPN)
    df = df.drop_duplicates(subset=["Order reference", "MPN"], keep="last")
    # финальный вид и порядок
    df = df[["MPN", "Replacem", "Quantity", "Totalsprice", "Order reference"]]
    return df.reset_index(drop=True)

# ---------- UI ----------
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
        "Алгоритм: ищем шапку (Artikuls/Daudz./Summa), строим окна колонок по X, "
        "а затем берём MPN/Daudz/Summa только из своих окон; Order — ближайший #1xxxxx."
    )
