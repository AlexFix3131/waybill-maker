# app.py
import io, re, statistics
from dataclasses import dataclass
from typing import List, Optional, Tuple

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

# ───────────────────────── Regex ─────────────────────────
# «Латвийский» MPN (11 цифр, допускаем ведущую C → удаляем)
RE_MPN_11   = re.compile(r"\b(?:C)?(\d{11})\b")

# MPN у Van Vliet: 2 цифры . 5 цифр - 4 цифры (пример: 81.36304-0019)
RE_MPN_VV   = re.compile(r"\b\d{2}\.\d{5}-\d{4}\b")

RE_INT      = re.compile(r"^\d{1,4}$")
RE_DEC      = re.compile(r"^\d{1,6}[.,]\d{2}$")
RE_MONEY    = re.compile(r"\d{1,3}(?:[ \u00A0]?\d{3})*[.,]\d{2}")

RE_HDR_ART  = re.compile(r"(?i)artik|artikul")
RE_HDR_QTY  = re.compile(r"(?i)daudz")
RE_HDR_SUM  = re.compile(r"(?i)summa|summ")

# Order / Reference
RE_ORDER = [
    re.compile(r"(?:^|\s)#\s*(1\d{5})(?:\s|$)"),
    re.compile(r"(?i)\border[_\-\s]*0*(1\d{5})"),
    re.compile(r"(?<![\d.,])(1\d{5})(?![\d.,])"),
]
RE_REFERENCE_LINE = re.compile(r"(?i)\breference[:\s]+(\d{5,})\b")

def to_float(s: str) -> float:
    return float(s.replace("\u00A0"," ").replace(" ","").replace(",","."))

def to_int(s: str) -> int:
    return int(round(to_float(s)))

def fmt_money_dot(s: Optional[str]) -> str:
    if not s:
        return "0.00"
    return f"{to_float(s):.2f}"

@dataclass
class Word:
    x0: float; y0: float; x1: float; y1: float; text: str

@dataclass
class Line:
    y: float; words: List[Word]; text: str

@dataclass
class Band:
    name: str; x_left: float; x_right: float

@dataclass
class OrderMark:
    x: float; y: float; value: str

# ───────────────────── helpers (общие) ─────────────────────
def load_words_per_page(pdf: bytes) -> List[List[Word]]:
    doc = fitz.open(stream=pdf, filetype="pdf")
    out: List[List[Word]] = []
    for p in doc:
        ws = [Word(w[0], w[1], w[2], w[3], w[4]) for w in p.get_text("words")]
        ws.sort(key=lambda w: (round(w.y0, 1), w.x0))
        out.append(ws)
    return out

def page_plain_texts(pdf: bytes) -> List[List[str]]:
    """Построчное «сырьё» для парсера таблиц формата Part/Description/Qty/Unit price/Sum."""
    doc = fitz.open(stream=pdf, filetype="pdf")
    pages: List[List[str]] = []
    for p in doc:
        txt = p.get_text("text")
        # нормализуем пробелы и неразрывные пробелы
        lines = [re.sub(r"[ \t\u00A0]+", " ", l).strip() for l in txt.splitlines()]
        lines = [l for l in lines if l]
        pages.append(lines)
    return pages

def group_lines(words: List[Word]) -> List[Line]:
    if not words: 
        return []
    heights = [w.y1 - w.y0 for w in words if (w.y1 - w.y0) > 0.2]
    h = statistics.median(heights) if heights else 8.0
    ytol = max(1.2, h * 0.65)

    res, cur, last = [], [], None
    for w in words:
        if last is None or abs(w.y0 - last) <= ytol:
            cur.append(w)
            last = w.y0 if last is None else (last + w.y0) / 2
        else:
            cur.sort(key=lambda t: t.x0)
            res.append(cur)
            cur = [w]
            last = w.y0
    if cur:
        cur.sort(key=lambda t: t.x0)
        res.append(cur)

    out: List[Line] = []
    for ln in res:
        y = statistics.fmean([w.y0 for w in ln])
        out.append(Line(y=y, words=ln, text=" ".join(w.text for w in ln)))
    out.sort(key=lambda L: L.y)
    return out

def detect_bands(lines: List[Line], words: List[Word]) -> List[Band]:
    # пытаемся найти заголовок «Artikuls / Daudz. / Summa»
    for L in lines[:80]:
        t = L.text
        if RE_HDR_ART.search(t) and RE_HDR_QTY.search(t) and RE_HDR_SUM.search(t):
            def center(pat):
                xs = [(w.x0 + w.x1) / 2 for w in L.words if pat.search(w.text)]
                return sum(xs) / len(xs) if xs else None
            cx_a = center(RE_HDR_ART); cx_q = center(RE_HDR_QTY); cx_s = center(RE_HDR_SUM)
            centers = [(n, c) for n, c in [("Artikuls", cx_a), ("Daudz.", cx_q), ("Summa", cx_s)] if c is not None]
            if len(centers) >= 2:
                centers.sort(key=lambda t: t[1])
                bands = []
                for i, (n, cx) in enumerate(centers):
                    left  = (centers[i-1][1] + cx) / 2 if i > 0 else cx - 90
                    right = (cx + centers[i+1][1]) / 2 if i < len(centers) - 1 else cx + 180
                    bands.append(Band(n, left, right))
                for b, nm in zip(sorted(bands, key=lambda b: b.x_left), ["Artikuls", "Daudz.", "Summa"]):
                    b.name = nm
                return bands
    # запасной вариант по ширине
    x_min = min(w.x0 for w in words); x_max = max(w.x1 for w in words); W = x_max - x_min
    return [
        Band("Artikuls", x_min - 10, x_min + 0.47 * W),
        Band("Daudz.",   x_min + 0.47 * W, x_min + 0.66 * W),
        Band("Summa",    x_min + 0.66 * W, x_max + 20),
    ]

def in_band(w: Word, b: Band) -> bool:
    cx = (w.x0 + w.x1) / 2
    return b.x_left <= cx <= b.x_right

# ─────────────────── Order detection (латвийский) ───────────────────
def order_from_text(txt: str) -> Optional[str]:
    for p in RE_ORDER:
        m = p.search(txt)
        if m: 
            return m.group(1)
    return None

def collect_order_marks(lines: List[Line]) -> List[OrderMark]:
    out: List[OrderMark] = []
    for L in lines:
        val = order_from_text(L.text)
        if val:
            xs = [(w.x0 + w.x1) / 2 for w in L.words if any(p.search(w.text) for p in RE_ORDER)]
            cx = statistics.median(xs) if xs else statistics.fmean([(w.x0 + w.x1) / 2 for w in L.words])
            out.append(OrderMark(cx, L.y, val))
    if not out:
        return []
    x_med = statistics.median([m.x for m in out])
    col = [m for m in out if abs(m.x - x_med) <= 42]
    out = col if len(col) >= max(3, len(out)//2) else out
    out.sort(key=lambda m: m.y)
    return out

# ───────────── фин. токены (латвийский), склейка денег ─────────────
def join_money_tokens(tokens: List[Word]) -> Optional[str]:
    if not tokens: 
        return None
    tokens.sort(key=lambda w: w.x0)
    groups = []; cur = [tokens[0]]
    for w in tokens[1:]:
        gap = w.x0 - cur[-1].x1
        if gap <= 8:
            cur.append(w)
        else:
            groups.append(cur); cur = [w]
    groups.append(cur)
    g = max(groups, key=lambda G: max(w.x1 for w in G))
    raw = "".join(w.text.replace("\u00A0","").replace(" ","") for w in g)
    if not re.search(r"[.,]\d{2}$", raw):
        raw = re.sub(r"(\d{2})$", r".\1", raw)
    return raw

def pick_total_for_line(line: Line, sum_band: Band) -> Optional[str]:
    cands = [w for w in line.words 
             if in_band(w, sum_band) and (
                 RE_MONEY.fullmatch(w.text) or RE_INT.fullmatch(w.text) or RE_DEC.fullmatch(w.text)
             )]
    if not cands:
        return None
    tok = join_money_tokens(cands)
    if not tok:
        return None
    # склейка «левая одиночная цифра» + «0xx,xx»
    lefts = [w for w in line.words if w.x1 <= cands[-1].x0 + 1 and (cands[-1].x0 - w.x1) <= 8]
    if lefts:
        lefts.sort(key=lambda w: w.x1, reverse=True)
        Lw = lefts[0]
        if re.fullmatch(r"[1-9]", Lw.text) and re.fullmatch(r"0\d{2}[.,]\d{2}", tok):
            tok = Lw.text + tok
    return tok

# ───────────────────────── Парсер №1 (латвийский) ─────────────────────────
def parse_latvian(pdf: bytes) -> pd.DataFrame:
    pages = load_words_per_page(pdf)
    rows = []
    for words in pages:
        if not words: 
            continue
        lines = group_lines(words)
        bands = detect_bands(lines, words)
        B = {b.name: b for b in bands}
        orders = collect_order_marks(lines)

        # кандидаты по MPN (11 цифр)
        cand_idx = []
        for i, L in enumerate(lines):
            if RE_MPN_11.search(L.text):
                cand_idx.append(i)

        for i in cand_idx:
            L = lines[i]
            m = RE_MPN_11.search(L.text)
            if not m:
                continue
            mpn = m.group(1)

            # qty из Daudz. (в той же строке, либо ±1 строка)
            qty_tok = None; best = (1e9, None)
            for d in [0, 1]:
                for sgn in (0, -1, 1):
                    j = i + sgn * d
                    if j < 0 or j >= len(lines): 
                        continue
                    for w in lines[j].words:
                        if in_band(w, B["Daudz."]) and (RE_INT.fullmatch(w.text) or RE_DEC.fullmatch(w.text)):
                            dy = abs(lines[j].y - L.y)
                            if dy < best[0]:
                                best = (dy, w.text)
                if best[1]:
                    break
            qty = to_int(best[1]) if best[1] else None

            # total (склейка токенов)
            bestT = (1e9, None)
            for d in [0, 1]:
                for sgn in (0, -1, 1):
                    j = i + sgn * d
                    if j < 0 or j >= len(lines): 
                        continue
                    tok = pick_total_for_line(lines[j], B["Summa"])
                    if tok:
                        dy = abs(lines[j].y - L.y)
                        if dy < bestT[0]:
                            bestT = (dy, tok)
                if bestT[1]:
                    break
            total = fmt_money_dot(bestT[1]) if bestT[1] else None

            # без пары qty+total строку отбрасываем
            if qty is None or total is None:
                continue

            order = nearest_order_above(orders, L.y)

            # если total == qty по числу — попробуем взять «самую правую» сумму в строке
            try:
                if abs(to_int(bestT[1]) - qty) == 0:
                    c = [w for w in L.words if in_band(w, B["Summa"]) and (
                        RE_MONEY.fullmatch(w.text) or RE_INT.fullmatch(w.text) or RE_DEC.fullmatch(w.text)
                    )]
                    if len(c) >= 2:
                        total = fmt_money_dot(join_money_tokens(c))
            except:
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
    df = pd.DataFrame(rows).drop_duplicates(keep="last")
    df = df[["MPN","Replacem","Quantity","Totalsprice","Order reference"]]
    df.reset_index(drop=True, inplace=True)
    return df

# ───────────────────────── Парсер №2 (Van Vliet: Part/Description/Qty/Unit/Sum) ─────────────────────────
def parse_van_vliet(pdf: bytes) -> pd.DataFrame:
    pages = page_plain_texts(pdf)

    # найдём Reference (Order)
    full_text = "\n".join(["\n".join(p) for p in pages])
    order = ""
    mref = RE_REFERENCE_LINE.search(full_text)
    if mref:
        order = mref.group(1)

    rows = []

    for page in pages:
        # ищем строку заголовков
        header_idx = -1
        for idx, line in enumerate(page[:50]):
            if (re.search(r"(?i)\bpart\b", line) and
                re.search(r"(?i)\bdescription\b", line) and
                re.search(r"(?i)\bqty\b", line) and
                re.search(r"(?i)\bsum\b", line)):
                header_idx = idx
                break
        if header_idx == -1:
            continue

        # идём вниз до подвала (Total / VAT / Invoice total …)
        for line in page[header_idx+1:]:
            if re.search(r"(?i)total|vat|invoice total", line):
                break

            # ожидаем, что строка начинается с MPN формата 81.36304-0019
            m_mpn = RE_MPN_VV.search(line)
            if not m_mpn:
                continue
            mpn = m_mpn.group(0)

            # выдёргиваем три последних числовых токена: qty, unit, sum
            nums = re.findall(r"\d+(?:[.,]\d+)?", line)
            # последние два — unit и sum, перед ними — qty (целое)
            if len(nums) < 3:
                continue
            sum_tok  = nums[-1]
            unit_tok = nums[-2]
            qty_tok  = nums[-3]

            # qty должен быть целым; если вдруг не целое — попробуем предыдущий токен
            if not re.fullmatch(r"\d+", qty_tok) and len(nums) >= 4 and re.fullmatch(r"\d+", nums[-4]):
                qty_tok = nums[-4]

            # финальный контроль
            if not re.fullmatch(r"\d+", qty_tok):
                continue

            qty = int(qty_tok)
            total = fmt_money_dot(sum_tok)

            rows.append({
                "MPN": mpn,
                "Replacem": "",
                "Quantity": qty,
                "Totalsprice": total,
                "Order reference": order
            })

    if not rows:
        return pd.DataFrame(columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])

    df = pd.DataFrame(rows).drop_duplicates(keep="last")
    df = df[["MPN","Replacem","Quantity","Totalsprice","Order reference"]]
    df.reset_index(drop=True, inplace=True)
    return df

# ───────────────────────── Комбайн ─────────────────────────
def parse_pdf(pdf: bytes) -> pd.DataFrame:
    """
    1) Пытаемся разобрать «латвийский» счет (Artikuls/Daudz./Summa, MPN=11 цифр).
    2) Если ничего не нашли — пробуем формат Van Vliet (Part/Description/Qty/Unit/Sum, MPN=XX.XXXXX-XXXX).
    """
    df = parse_latvian(pdf)
    if not len(df):
        df = parse_van_vliet(pdf)
    return df

# ───────────────────────── UI ─────────────────────────
pdf_file = st.file_uploader("Загрузить PDF-счёт", type=["pdf"])
tpl_file = st.file_uploader("Шаблон Excel (необязательно)", type=["xlsx"])

if pdf_file:
    df = parse_pdf(pdf_file.read())

    if df.empty:
        st.warning("Новый/экзотический формат счёта. Показываю первые строки текста для диагностики.")
        # выводим начало текста, чтобы тебе было проще прислать пример
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        doc.close()
        txt_pages = page_plain_texts(pdf_file.read())
        st.code("\n".join("\n".join(p[:60]) for p in txt_pages[:2]))
    else:
        st.subheader("Предпросмотр")
        st.dataframe(df, use_container_width=True)

    if not df.empty and st.button("Скачать Excel"):
        if tpl_file:
            wb = load_workbook(tpl_file)
            ws = wb.active
            if ws.max_row <= 1:
                ws.append(["MPN","Replacem","Quantity","Totalsprice","Order reference"])
        else:
            wb = Workbook()
            ws = wb.active
            ws.append(["MPN","Replacem","Quantity","Totalsprice","Order reference"])

        for r in df.itertuples(index=False):
            ws.append(list(r))

        bio = io.BytesIO()
        wb.save(bio)
        st.download_button(
            "Скачать waybill.xlsx",
            data=bio.getvalue(),
            file_name="waybill.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info(
        "Логика: 1) фильтруем MPN по формату (11 цифр **или** XX.XXXXX-XXXX); "
        "2) ищем Qty (целое) и Sum (склейка денег) в тех же/соседних строках; "
        "3) Order — из Reference в шапке либо ближайший сверху."
    )
