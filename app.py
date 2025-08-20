# app.py

import io, re, fitz, streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook

# ==========================
# helpers
# ==========================

def clean_num(s: str) -> float | None:
    """
    Преобразует '1.027,07' / '4 106,79' / '545,7' / '31,34' / '6349.20' -> float
    Возвращает None, если не похоже на число.
    """
    if s is None:
        return None
    s = str(s).strip()
    # убираем пробелы тыс.
    s = s.replace(' ', '')
    # если есть и '.' и ',' — предполагаем . = thousands, , = decimal
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    else:
        # один разделитель — если запятая, делаем точку
        s = s.replace(',', '.')
    try:
        return float(s)
    except Exception:
        return None

def clean_int(s: str) -> int | None:
    v = clean_num(s)
    try:
        if v is None:
            return None
        # кол-во бывает 26.00 -> 26
        return int(round(v))
    except Exception:
        return None

def strip_leading_c(mpn: str) -> str:
    """
    "C81.36400-6007" -> "81.36400-6007" (правило "убрать C")
    """
    if mpn and mpn.upper().startswith('C') and len(mpn) > 1 and mpn[1].isdigit():
        return mpn[1:]
    return mpn

def dedupe_rows(rows):
    seen = set()
    out = []
    for r in rows:
        key = (r['MPN'], r['Quantity'], r['Totalsprice'], r['Order reference'])
        if key in seen: 
            continue
        seen.add(key)
        out.append(r)
    return out

def to_df(rows):
    rows = dedupe_rows(rows)
    if not rows:
        return pd.DataFrame(columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])
    return pd.DataFrame(rows, columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])

def pdf_text_lines(pdf_bytes: bytes) -> list[str]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    lines = []
    for p in doc:
        txt = p.get_text()
        # разбиваем именно по строкам
        lines.extend([l.rstrip() for l in txt.splitlines()])
    doc.close()
    return lines

# ==========================
# supplier detection
# ==========================

def is_japafrica(lines: list[str]) -> bool:
    join = ' '.join(lines)
    return ('JAPAFRICA MOBILITY SOLUTIONS' in join) or ('FACTURA' in join and 'Backorder_' in join)

def is_vanvliet(lines: list[str]) -> bool:
    join = ' '.join(lines)
    return ('Van Vliet TechSupport' in join) or ('INVOICE' in join and 'Reference:' in join and 'Qty' in join and 'Unit price' in join)

# ==========================
# JAPAFRICA parser
# ==========================

def parse_order_japafrica(lines: list[str]) -> str | None:
    join = ' '.join(lines)
    m = re.search(r'Backorder[_\-\s]*?(\d{4,})', join, re.IGNORECASE)
    if m:
        return m.group(1)
    # запасной: иногда только Enc/Req (не идеально, но лучше, чем ничего)
    m = re.search(r'\bEnc/Req\.\s*([0-9_]+)', join)
    if m:
        digits = ''.join(ch for ch in m.group(1) if ch.isdigit())
        if digits:
            return digits
    return None

def parse_lines_japafrica(lines: list[str], order_no: str | None) -> list[dict]:
    """
    Пытаемся поймать строки вида:
    C81.36400-6007 REPAIR KIT GASKET  1B4  244.20  26.00  UNID  6349.20
    """
    rows = []
    rx = re.compile(
        r'\b([A-Z0-9][A-Z0-9.\-]+)\s+.+?\s+[A-Z0-9]{2,}\s+(\d{1,3}[.,]\d{2})\s+(\d{1,4}(?:[.,]\d{2})?)\s+(?:UNID|UN|PCS)?\s+(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\b'
    )
    for ln in lines:
        m = rx.search(ln)
        if not m:
            continue
        mpn = strip_leading_c(m.group(1))
        unit_price = clean_num(m.group(2))   # не нужен, но валидируем
        qty = clean_int(m.group(3))
        total = clean_num(m.group(4))
        if not mpn or qty is None or total is None:
            continue
        rows.append({
            "MPN": mpn,
            "Replacem": "",
            "Quantity": qty,
            "Totalsprice": total,
            "Order reference": order_no or ""
        })
    return rows

def parse_blocks_japafrica(pdf_bytes: bytes, order_no: str | None) -> list[dict]:
    """
    Фолбэк по блокам (координаты): ищем MPN в одном блоке и рядом справа три числа.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    rows = []
    for p in doc:
        blocks = p.get_text("blocks")  # [(x0,y0,x1,y1, "text", block_no, block_type, ...)]
        # Сортировка по y, потом x
        blocks.sort(key=lambda b: (round(b[1],1), round(b[0],1)))
        for b in blocks:
            txt = b[4].strip()
            # MPN-подобная строка
            if not re.search(r'\b[A-Z0-9][A-Z0-9.\-]{6,}\b', txt):
                continue
            mpn_match = re.search(r'\b([A-Z0-9][A-Z0-9.\-]{6,})\b', txt)
            if not mpn_match:
                continue
            raw_mpn = strip_leading_c(mpn_match.group(1))
            y_mid = (b[1]+b[3])/2
            # ищем 3 числовых блока на той же строке правее
            numeric_candidates = []
            for bb in blocks:
                if bb[0] <= b[2]:   # правее
                    continue
                y_mid2 = (bb[1]+bb[3])/2
                if abs(y_mid2 - y_mid) > 4.0:
                    continue
                text2 = bb[4].strip()
                nums = re.findall(r'\d{1,3}(?:[.,]\d{3})*[.,]\d{2}', text2)
                if nums:
                    numeric_candidates.extend(nums)
            # пытаемся вычленить qty и total
            # qty иногда идёт как 26.00, а total как 6349.20; берём самое крупное за total
            parsed = [clean_num(n) for n in numeric_candidates]
            parsed = [x for x in parsed if x is not None]
            if len(parsed) < 2:
                continue
            total = max(parsed)
            # qty — ближайшее число в районе 1..5000, но обычно 1..1000
            qty_candidates = [x for x in parsed if 0 < x <= 10000 and abs(x-total) > 1e-6]
            qty_val = None
            if qty_candidates:
                # предпочитаем целые
                ints = [int(round(v)) for v in qty_candidates if abs(v-round(v)) < 1e-6]
                if ints:
                    qty_val = ints[0]
                else:
                    qty_val = int(round(qty_candidates[0]))
            if qty_val is None:
                continue
            rows.append({
                "MPN": raw_mpn,
                "Replacem": "",
                "Quantity": qty_val,
                "Totalsprice": total,
                "Order reference": order_no or ""
            })
    doc.close()
    return rows

def parse_japafrica(pdf_bytes: bytes, lines: list[str]) -> pd.DataFrame:
    order_no = parse_order_japafrica(lines)
    rows = parse_lines_japafrica(lines, order_no)
    if not rows:
        rows = parse_blocks_japafrica(pdf_bytes, order_no)
    return to_df(rows)

# ==========================
# VAN VLIET parser
# ==========================

def parse_order_vanvliet(lines: list[str]) -> str | None:
    join = ' '.join(lines)
    m = re.search(r'\bReference:\s*(\d{4,})', join, re.IGNORECASE)
    if m:
        return m.group(1)
    return None

def parse_vanvliet(lines: list[str]) -> pd.DataFrame:
    order_no = parse_order_vanvliet(lines)
    rows = []
    # строка вида:
    # 06.01494-6735 Hex shoulder stud ... 8 14,74 117,94
    rx = re.compile(
        r'^\s*([0-9]{2}\.[0-9]{5}-[0-9]{4}|[0-9]{2}\.[0-9]{5}-[0-9]{3}|[0-9]{2}\.[0-9]{5}-[0-9]{1,4}|[0-9]{2}\.[0-9]{5})\s+.+?\s+(\d{1,6})\s+(\d{1,3}[.,]\d{2})\s+(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\s*$'
    )
    for ln in lines:
        m = rx.match(ln)
        if not m:
            continue
        mpn = m.group(1)
        qty = clean_int(m.group(2))
        total = clean_num(m.group(4))
        if not mpn or qty is None or total is None:
            continue
        rows.append({
            "MPN": mpn,
            "Replacem": "",
            "Quantity": qty,
            "Totalsprice": total,
            "Order reference": order_no or ""
        })
    return to_df(rows)

# ==========================
# generic / UI
# ==========================

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

pdf = st.file_uploader("Загрузить PDF‑счет", type=["pdf"])
tpl = st.file_uploader("Шаблон Excel (необязательно)", type=["xlsx"])

if pdf:
    pdf_bytes = pdf.read()
    lines = pdf_text_lines(pdf_bytes)

    # определяем поставщика
    if is_vanvliet(lines):
        st.caption("Обнаружен формат: **Van Vliet TechSupport**")
        df = parse_vanvliet(lines)
    elif is_japafrica(lines):
        st.caption("Обнаружен формат: **JAPAFRICA**")
        df = parse_japafrica(pdf_bytes, lines)
    else:
        st.warning("Новый/неизвестный формат счёта. Показываю первые строки текста для диагностики.")
        st.code("\n".join(lines[:120]))
        df = pd.DataFrame(columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])

    # финальный вид и редактор
    st.subheader("Предпросмотр")
    st.dataframe(df, use_container_width=True)

    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)

    # кнопка скачать
    if st.button("Скачать Excel"):
        # подготавливаем книгу
        if tpl:
            wb = load_workbook(tpl); ws = wb.active
        else:
            wb = Workbook(); ws = wb.active
            ws["A1"]="MPN"; ws["B1"]="Replacem"; ws["C1"]="Quantity"; ws["D1"]="Totalsprice"; ws["E1"]="Order reference"
        # чистим старые данные (2..2000)
        for r in range(2, 2001):
            for c in range(1, 6):
                ws.cell(row=r, column=c).value = None
        # записываем
        for i, row in enumerate(edited.itertuples(index=False), start=2):
            ws.cell(i, 1, row[0])                     # MPN
            ws.cell(i, 2, row[1])                     # Replacem
            ws.cell(i, 3, int(row[2]) if row[2] != "" and pd.notna(row[2]) else None)  # qty
            ws.cell(i, 4, float(row[3]) if row[3] != "" and pd.notna(row[3]) else None) # total
            ws.cell(i, 5, row[4])                     # order

        bio = io.BytesIO()
        wb.save(bio)
        st.download_button(
            "Скачать waybill.xlsx",
            data=bio.getvalue(),
            file_name="waybill.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Загрузи PDF счёт — я распарсю и заполню Excel по шаблону.")
