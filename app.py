# app.py
import io, re, statistics
from dataclasses import dataclass
from typing import List, Optional

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

# ---------- Regex ----------
# NEW: любые 11 цифр (с опц. ведущей C, которую удаляем)
RE_MPN   = re.compile(r"\b(?:C)?(\d{11})\b")
RE_MONEY = re.compile(r"\d{1,3}(?:[ \u00A0]?\d{3})*[.,]\d{2}")
RE_DEC   = re.compile(r"^\d{1,6}[.,]\d{2}$")

RE_HDR_ART = re.compile(r"(?i)artik|artikul")
RE_HDR_QTY = re.compile(r"(?i)daudz")
RE_HDR_SUM = re.compile(r"(?i)summa|summ")

RE_ORDER = [
    re.compile(r"(?:^|\s)#\s*(1\d{5})(?:\s|$)"),
    re.compile(r"(?i)\border[_\-\s]*0*(1\d{5})"),
    re.compile(r"(?<![\d.,])(1\d{5})(?![\d.,])"),
]

def to_float(s: str) -> float:
    return float(s.replace("\u00A0"," ").replace(" ","").replace(",","."))
def to_int(s: str) -> int:
    return int(round(to_float(s)))

# NEW: теперь всегда отдаем строку с ТОЧКОЙ
def fmt_money_dot(s: Optional[str]) -> str:
    if not s: return "0.00"
    try:
        return f"{to_float(s):.2f}"
    except Exception:
        # если вдруг не разобралась — последний шанс
        return s.replace(",", ".")

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

# ---------- PDF helpers ----------
def load_words_per_page(pdf: bytes) -> List[List[Word]]:
    doc = fitz.open(stream=pdf, filetype="pdf")
    out=[]
    for p in doc:
        ws=[Word(w[0],w[1],w[2],w[3],w[4]) for w in p.get_text("words")]
        ws.sort(key=lambda w:(round(w.y0,1), w.x0))
        out.append(ws)
    return out

def group_lines(words: List[Word]) -> List[Line]:
    if not words: return []
    heights=[w.y1-w.y0 for w in words if (w.y1-w.y0)>0.2]
    h = statistics.median(heights) if heights else 8.0
    ytol=max(1.2, h*0.65)

    cur=[]; last=None; res=[]
    for w in words:
        if last is None or abs(w.y0-last)<=ytol:
            cur.append(w); last = w.y0 if last is None else (last+w.y0)/2
        else:
            cur.sort(key=lambda t:t.x0)
            res.append(cur); cur=[w]; last=w.y0
    if cur: cur.sort(key=lambda t:t.x0); res.append(cur)

    out=[]
    for ln in res:
        y=statistics.fmean([w.y0 for w in ln])
        out.append(Line(y=y, words=ln, text=" ".join(w.text for w in ln)))
    out.sort(key=lambda L:L.y)
    return out

def detect_bands(lines: List[Line], words: List[Word]) -> List[Band]:
    for L in lines[:80]:
        t=L.text
        if RE_HDR_ART.search(t) and RE_HDR_QTY.search(t) and RE_HDR_SUM.search(t):
            def center(pat):
                xs=[(w.x0+w.x1)/2 for w in L.words if pat.search(w.text)]
                return sum(xs)/len(xs) if xs else None
            cx_a=center(RE_HDR_ART); cx_q=center(RE_HDR_QTY); cx_s=center(RE_HDR_SUM)
            centers=[(n,c) for n,c in [("Artikuls",cx_a),("Daudz.",cx_q),("Summa",cx_s)] if c is not None]
            if len(centers)>=2:
                centers.sort(key=lambda t:t[1])
                bands=[]
                for i,(n,cx) in enumerate(centers):
                    left  = (centers[i-1][1]+cx)/2 if i>0 else cx-90
                    right = (cx+centers[i+1][1])/2 if i<len(centers)-1 else cx+180
                    bands.append(Band(n,left,right))
                for b,nm in zip(sorted(bands,key=lambda b:b.x_left),["Artikuls","Daudz.","Summa"]):
                    b.name=nm
                return bands
    # грубо по ширине
    x_min=min(w.x0 for w in words); x_max=max(w.x1 for w in words); W=x_max-x_min
    return [
        Band("Artikuls", x_min-10, x_min+0.47*W),
        Band("Daudz.",   x_min+0.47*W, x_min+0.66*W),
        Band("Summa",    x_min+0.66*W, x_max+20),
    ]

def in_band(w: Word, b: Band) -> bool:
    cx=(w.x0+w.x1)/2
    return b.x_left<=cx<=b.x_right

# ---------- Order ----------
def order_from_text(txt: str) -> Optional[str]:
    for p in RE_ORDER:
        m=p.search(txt)
        if m: return m.group(1)
    return None

def collect_order_marks(lines: List[Line]) -> List[OrderMark]:
    out=[]
    for L in lines:
        val=order_from_text(L.text)
        if val:
            xs=[(w.x0+w.x1)/2 for w in L.words if any(p.search(w.text) for p in RE_ORDER)]
            cx=statistics.median(xs) if xs else statistics.fmean([(w.x0+w.x1)/2 for w in L.words])
            out.append(OrderMark(cx, L.y, val))
    if not out: return []
    x_med=statistics.median([m.x for m in out])
    col=[m for m in out if abs(m.x-x_med)<=42]
    out = col if len(col)>=max(3, len(out)//2) else out
    out.sort(key=lambda m:m.y)
    return out

def order_for_line_y(marks: List[OrderMark], y: float) -> str:
    prev=[m for m in marks if m.y<=y+2]
    if prev: return prev[-1].value
    if not marks: return ""
    best=min(marks, key=lambda m: abs(m.y-y))
    return best.value if abs(best.y-y)<=30 else ""

# ---------- Core ----------
def parse_pdf(pdf: bytes) -> pd.DataFrame:
    pages=load_words_per_page(pdf)
    rows=[]
    for words in pages:
        if not words: continue
        lines=group_lines(words)
        bands=detect_bands(lines, words)
        b = {b.name:b for b in bands}

        order_marks=collect_order_marks(lines)

        mpn_idx=[]
        for i,L in enumerate(lines):
            if RE_MPN.search(L.text):
                mpn_idx.append(i)

        for i in mpn_idx:
            L=lines[i]
            m=RE_MPN.search(L.text)
            if not m: continue
            mpn=m.group(1)  # уже без 'C'

            order = order_for_line_y(order_marks, L.y)

            # qty (Daudz.) — ближайший по Y (смотрим до ±2 строк)
            qty_tok=None; best=(1e9,None)
            for d in [0,1,2]:
                for sgn in (0,-1,1):
                    j=i+sgn*d
                    if j<0 or j>=len(lines): continue
                    for w in lines[j].words:
                        if in_band(w, b["Daudz."]) and RE_DEC.match(w.text):
                            dy=abs(lines[j].y - L.y)
                            if dy<best[0]:
                                best=(dy,w.text)
                if best[1]: break
            qty = to_int(best[1]) if best[1] else 0

            # total (Summa) — правая сумма + склейка "левое число" + "0xx,xx"
            def pick_total(line: Line) -> Optional[str]:
                c=[w for w in line.words if in_band(w,b["Summa"]) and RE_MONEY.fullmatch(w.text)]
                if not c: return None
                c.sort(key=lambda w:max(w.x0,w.x1))
                t=c[-1]
                tok=t.text
                # склейка с цифрой слева
                lefts=[w for w in line.words if w.x1<=t.x0+1 and (t.x0-w.x1)<=8]
                if lefts:
                    lefts.sort(key=lambda w:w.x1, reverse=True)
                    lw=lefts[0]
                    if re.fullmatch(r"[1-9]", lw.text) and re.fullmatch(r"0\d{2}[.,]\d{2}", tok):
                        tok = lw.text + tok
                return tok

            total_tok=None; bestT=(1e9,None)
            for d in [0,1,2]:
                for sgn in (0,-1,1):
                    j=i+sgn*d
                    if j<0 or j>=len(lines): continue
                    tok=pick_total(lines[j])
                    if tok:
                        dy=abs(lines[j].y - L.y)
                        if dy<bestT[0]:
                            bestT=(dy,tok)
                if bestT[1]: break
            total = fmt_money_dot(bestT[1])

            # если total == qty (по значению) и есть ещё суммы — возьми самую правую из строки
            try:
                if bestT[1] and abs(to_int(bestT[1]) - qty) == 0:
                    c=[w for w in L.words if in_band(w,b["Summa"]) and RE_MONEY.fullmatch(w.text)]
                    c.sort(key=lambda w:max(w.x0,w.x1))
                    if len(c)>=2:
                        total = fmt_money_dot(c[-1].text)
            except: pass

            rows.append({
                "MPN": mpn,
                "Replacem": "",
                "Quantity": qty,
                "Totalsprice": total,         # уже с точкой
                "Order reference": order
            })

    if not rows:
        return pd.DataFrame(columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])
    df=pd.DataFrame(rows).drop_duplicates(keep="last")
    df=df[["MPN","Replacem","Quantity","Totalsprice","Order reference"]]
    df.reset_index(drop=True, inplace=True)
    return df

# ---------- UI ----------
pdf_file = st.file_uploader("Загрузить PDF-счёт", type=["pdf"])
tpl_file = st.file_uploader("Шаблон Excel (необязательно)", type=["xlsx"])

if pdf_file:
    df=parse_pdf(pdf_file.read())
    st.subheader("Предпросмотр")
    st.dataframe(df, use_container_width=True)

    if st.button("Скачать Excel"):
        if tpl_file:
            wb=load_workbook(tpl_file); ws=wb.active
        else:
            wb=Workbook(); ws=wb.active
            ws.append(["MPN","Replacem","Quantity","Totalsprice","Order reference"])
        for r in df.itertuples(index=False):
            ws.append(list(r))
        bio=io.BytesIO(); wb.save(bio)
        st.download_button("Скачать waybill.xlsx",
                           data=bio.getvalue(),
                           file_name="waybill.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Парсер: MPN = любые 11 цифр (без ведущей C); Qty — колонка Daudz.; Total — правая сумма из Summa со склейкой; Order — ближайший сверху.")
