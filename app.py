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

# ───────────────────────── Common regex ─────────────────────────
# Materom (латышский макет): MPN = 11 цифр, может быть ведущая 'C' — её убираем
RE_MPN_11   = re.compile(r"\b(?:C)?(\d{11})\b")
RE_INT      = re.compile(r"^\d{1,4}$")
RE_DEC      = re.compile(r"^\d{1,6}[.,]\d{2}$")
RE_MONEY    = re.compile(r"\d{1,3}(?:[ \u00A0]?\d{3})*[.,]\d{2}")

RE_HDR_ART  = re.compile(r"(?i)artik|artikul")
RE_HDR_QTY  = re.compile(r"(?i)daudz")
RE_HDR_SUM  = re.compile(r"(?i)summa|summ")

# Van Vliet: Part формат вида 06.01494-6735 / 81.36304-0013 → после чистки остаётся 11 цифр
RE_VV_PART  = re.compile(r"\b\d{2}\.\d{5}-\d{4}\b")   # как в счёте Van Vliet:contentReference[oaicite:2]{index=2}
RE_VV_HDR   = re.compile(r"(?i)\bPart\b.*\bDescription\b.*\bQty\b.*\bSum\s*\(EUR\)\b")  # заголовок:contentReference[oaicite:3]{index=3}
RE_VV_REF   = re.compile(r"(?i)\bReference\s*:\s*(\d+)\b")  # order для Van Vliet:contentReference[oaicite:4]{index=4}

# универсальные конвертеры
def to_float(s: str) -> float:
    return float(s.replace("\u00A0"," ").replace(" ","").replace(".","").replace(",","."))
def to_int(s: str) -> int:
    return int(round(to_float(s)))
def fmt_money_dot(s: Optional[str]) -> str:
    if not s: return "0.00"
    return f"{to_float(s):.2f}"

# ───────────────────────── Geometry models ─────────────────────────
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

# ───────────────────────── Page helpers ─────────────────────────
def load_words_per_page(pdf: bytes) -> List[List[Word]]:
    doc = fitz.open(stream=pdf, filetype="pdf")
    out=[]
    for p in doc:
        ws=[Word(w[0],w[1],w[2],w[3],w[4]) for w in p.get_text("words")]
        ws.sort(key=lambda w:(round(w.y0,1), w.x0))
        out.append(ws)
    return out

def join_lines(words: List[Word]) -> List[Line]:
    if not words: return []
    heights=[w.y1-w.y0 for w in words if (w.y1-w.y0)>0.2]
    h = statistics.median(heights) if heights else 8.0
    ytol=max(1.2, h*0.65)

    res=[]; cur=[]; last=None
    for w in words:
        if last is None or abs(w.y0-last)<=ytol:
            cur.append(w); last = w.y0 if last is None else (last+w.y0)/2
        else:
            cur.sort(key=lambda t:t.x0); res.append(cur); cur=[w]; last=w.y0
    if cur: cur.sort(key=lambda t:t.x0); res.append(cur)

    out=[]
    for ln in res:
        y=statistics.fmean([w.y0 for w in ln])
        out.append(Line(y=y, words=ln, text=" ".join(w.text for w in ln)))
    out.sort(key=lambda L:L.y)
    return out

# ───────────────────────── Materom parser ─────────────────────────
RE_ORDER_M = [
    re.compile(r"(?:^|\s)#\s*(1\d{5})(?:\s|$)"),
    re.compile(r"(?i)\border[_\-\s]*0*(1\d{5})"),
    re.compile(r"(?<![\d.,])(1\d{5})(?![\d.,])"),
]

def order_from_text(txt: str) -> Optional[str]:
    for p in RE_ORDER_M:
        m=p.search(txt)
        if m: return m.group(1)
    return None

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
    # запасной вариант по ширине
    x_min=min(w.x0 for w in words); x_max=max(w.x1 for w in words); W=x_max-x_min
    return [
        Band("Artikuls", x_min-10, x_min+0.47*W),
        Band("Daudz.",   x_min+0.47*W, x_min+0.66*W),
        Band("Summa",    x_min+0.66*W, x_max+20),
    ]

def in_band(w: Word, b: Band) -> bool:
    cx=(w.x0+w.x1)/2
    return b.x_left<=cx<=b.x_right

def collect_order_marks(lines: List[Line]) -> List[OrderMark]:
    out=[]
    for L in lines:
        val=order_from_text(L.text)
        if val:
            xs=[(w.x0+w.x1)/2 for w in L.words if any(p.search(w.text) for p in RE_ORDER_M)]
            cx=statistics.median(xs) if xs else statistics.fmean([(w.x0+w.x1)/2 for w in L.words])
            out.append(OrderMark(cx, L.y, val))
    if not out: return []
    x_med=statistics.median([m.x for m in out])
    col=[m for m in out if abs(m.x-x_med)<=42]
    out = col if len(col)>=max(3, len(out)//2) else out
    out.sort(key=lambda m:m.y)
    return out

def nearest_order_above(marks: List[OrderMark], y: float) -> str:
    prev=[m for m in marks if m.y<=y+2]
    if prev: return prev[-1].value
    if not marks: return ""
    best=min(marks, key=lambda m: abs(m.y-y))
    return best.value if abs(best.y-y)<=30 else ""

def join_money_tokens(tokens: List[Word]) -> Optional[str]:
    if not tokens: return None
    tokens.sort(key=lambda w: w.x0)
    groups=[]; cur=[tokens[0]]
    for w in tokens[1:]:
        gap = w.x0 - cur[-1].x1
        if gap <= 8:
            cur.append(w)
        else:
            groups.append(cur); cur=[w]
    groups.append(cur)
    g = max(groups, key=lambda G: max(w.x1 for w in G))
    raw = "".join(w.text.replace("\u00A0","").replace(" ","") for w in g)
    if not re.search(r"[.,]\d{2}$", raw):
        raw = re.sub(r"(\d{2})$", r".\1", raw)
    return raw

def pick_total_for_line(line: Line, sum_band: Band) -> Optional[str]:
    cands=[w for w in line.words if in_band(w,sum_band) and (RE_MONEY.fullmatch(w.text) or RE_INT.fullmatch(w.text) or RE_DEC.fullmatch(w.text))]
    if not cands: return None
    tok = join_money_tokens(cands)
    if not tok: return None
    # склейка «1» + «0xx,xx»
    lefts=[w for w in line.words if w.x1<=cands[-1].x0+1 and (cands[-1].x0-w.x1)<=8]
    if lefts:
        lefts.sort(key=lambda w:w.x1, reverse=True)
        Lw=lefts[0]
        if re.fullmatch(r"[1-9]", Lw.text) and re.fullmatch(r"0\d{2}[.,]\d{2}", tok):
            tok = Lw.text + tok
    return tok

def parse_materom(pdf: bytes) -> pd.DataFrame:
    pages=load_words_per_page(pdf)
    rows=[]
    for words in pages:
        if not words: continue
        lines=join_lines(words)
        bands=detect_bands(lines, words)
        B = {b.name:b for b in bands}
        orders=collect_order_marks(lines)

        cand_idx=[]
        for i,L in enumerate(lines):
            if RE_MPN_11.search(L.text):
                cand_idx.append(i)

        for i in cand_idx:
            L=lines[i]
            m=RE_MPN_11.search(L.text)
            if not m: continue
            mpn=m.group(1)

            # qty
            qty_tok=None; best=(1e9,None)
            for d in [0,1]:
                for sgn in (0,-1,1):
                    j=i+sgn*d
                    if j<0 or j>=len(lines): continue
                    for w in lines[j].words:
                        if in_band(w, B["Daudz."]) and (RE_INT.fullmatch(w.text) or RE_DEC.fullmatch(w.text)):
                            dy=abs(lines[j].y - L.y)
                            if dy<best[0]:
                                best=(dy,w.text)
                if best[1]: break
            qty = to_int(best[1]) if best[1] else None

            # total
            bestT=(1e9,None)
            for d in [0,1]:
                for sgn in (0,-1,1):
                    j=i+sgn*d
                    if j<0 or j>=len(lines): continue
                    tok=pick_total_for_line(lines[j], B["Summa"])
                    if tok:
                        dy=abs(lines[j].y - L.y)
                        if dy<bestT[0]:
                            bestT=(dy,tok)
                if bestT[1]: break
            total = fmt_money_dot(bestT[1]) if bestT[1] else None

            if qty is None or total is None:
                continue

            order = nearest_order_above(orders, L.y)

            try:
                if abs(to_int(bestT[1]) - qty) == 0:
                    c=[w for w in L.words if in_band(w,B["Summa"]) and (RE_MONEY.fullmatch(w.text) or RE_INT.fullmatch(w.text) or RE_DEC.fullmatch(w.text))]
                    if len(c)>=2:
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
    df=pd.DataFrame(rows).drop_duplicates(keep="last")
    df=df[["MPN","Replacem","Quantity","Totalsprice","Order reference"]]
    df.reset_index(drop=True, inplace=True)
    return df

# ───────────────────────── Van Vliet parser ─────────────────────────
def looks_like_van_vliet(all_text: str) -> bool:
    # В шапке встречается таблица Part|Description|Qty|Unit price|Sum (EUR):contentReference[oaicite:5]{index=5}
    return bool(RE_VV_HDR.search(all_text)) or "VAN VLIET" in all_text.upper()

def clean_vv_mpn(token: str) -> Optional[str]:
    # "06.01494-6735" -> "06014946735" (11 цифр)
    digits = re.sub(r"\D","", token)
    return digits if len(digits)==11 else None

def parse_van_vliet(pdf: bytes) -> pd.DataFrame:
    doc = fitz.open(stream=pdf, filetype="pdf")
    pages_words = [ [Word(*w[:5]) for w in p.get_text("words")] for p in doc ]
    all_text = "\n".join(p.get_text("text") for p in doc)

    # единый order nr из поля Reference: NNNNNN:contentReference[oaicite:6]{index=6}
    order = ""
    m = RE_VV_REF.search(all_text)
    if m: order = m.group(1)

    rows=[]
    for page_idx, words in enumerate(pages_words):
        if not words: continue
        lines = join_lines(words)

        for L in lines:
            # строка с Part
            mpart = RE_VV_PART.search(L.text)
            if not mpart:
                continue
            mpn_raw = mpart.group(0)
            mpn = clean_vv_mpn(mpn_raw)
            if not mpn:
                continue

            # рядом в строке должны быть Qty и Sum (EUR)
            # Qty — целое; Sum — деньги (с запятой)
            qty = None
            total = None

            # Берём токены в этой линии и в линии ниже — иногда цена уходит на строку
            neighbours = [L]
            # добавим следующую строку, если она очень близко по y
            li = lines.index(L)
            if li+1 < len(lines) and abs(lines[li+1].y - L.y) < 12:
                neighbours.append(lines[li+1])

            # qty: самое правое целое число рядом с part
            qty_candidates=[]
            sum_candidates=[]
            for ln in neighbours:
                for w in ln.words:
                    if RE_INT.fullmatch(w.text):
                        qty_candidates.append((w.x0, w.text))
                    elif RE_MONEY.fullmatch(w.text) or RE_DEC.fullmatch(w.text):
                        sum_candidates.append((w.x0, w.text))

            if qty_candidates:
                qty = to_int(sorted(qty_candidates, key=lambda t:t[0])[-1][1])
            if sum_candidates:
                # берём крайний правый токен как сумму (в этой разметке Sum(EUR) справа)
                total = fmt_money_dot(sorted(sum_candidates, key=lambda t:t[0])[-1][1])

            if qty is None or total is None:
                # если что-то не нашли в текущих строках — попробуем ещё одну вниз
                if li+2 < len(lines) and qty is None:
                    for w in lines[li+2].words:
                        if RE_INT.fullmatch(w.text):
                            qty = to_int(w.text)
                            break
                if li+2 < len(lines) and total is None:
                    mny=[w.text for w in lines[li+2].words if RE_MONEY.fullmatch(w.text) or RE_DEC.fullmatch(w.text)]
                    if mny:
                        total = fmt_money_dot(mny[-1])

            if qty is None or total is None:
                continue

            rows.append({
                "MPN": mpn,
                "Replacem": "",
                "Quantity": qty,
                "Totalsprice": total,
                "Order reference": order,
            })

    if not rows:
        return pd.DataFrame(columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])
    df = pd.DataFrame(rows).drop_duplicates(keep="last")
    df=df[["MPN","Replacem","Quantity","Totalsprice","Order reference"]]
    df.reset_index(drop=True, inplace=True)
    return df

# ───────────────────────── Dispatcher ─────────────────────────
def parse_pdf(pdf: bytes) -> pd.DataFrame:
    # Быстрый детектор макета
    try:
        txt = fitz.open(stream=pdf, filetype="pdf")[0].get_text("text")
    except:
        txt = ""
    if looks_like_van_vliet(txt):
        return parse_van_vliet(pdf)
    else:
        return parse_materom(pdf)

# ───────────────────────── UI ─────────────────────────
pdf_file = st.file_uploader("Загрузить PDF-счёт", type=["pdf"])
tpl_file = st.file_uploader("Шаблон Excel (необязательно)", type=["xlsx"])

if pdf_file:
    df=parse_pdf(pdf_file.read())
    # подсказка если совсем пусто — показать первые строки текста для дебага
    if df.empty:
        st.warning("Новый/неопознанный формат счёта. Показываю первые строки текста для диагностики.")
        try:
            doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        except:
            doc = fitz.open(stream=pdf_file.getvalue(), filetype="pdf")
        t = doc[0].get_text("text").splitlines()[:40]
        st.code("\n".join(t))
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
    st.info("Парсеры: ① Materom (латышский макет Artikuls/Daudz./Summa + order по «Order_...»), "
            "② Van Vliet (таблица Part|Description|Qty|Sum (EUR), order берём из поля Reference).")
