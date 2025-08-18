import streamlit as st, re, io, pandas as pd
from PyPDF2 import PdfReader
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from openpyxl import load_workbook, Workbook
from utils import load_profiles, cleanse_mpn

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

# ---------- правила (yaml -> fallback) ----------
def load_rules_safe():
    try:
        profiles = load_profiles("supplier_profiles.yaml")
        return profiles.get("default", {})
    except Exception:
        return {
            "remove_leading_C_in_mpn": True,
            "materom_mpn_before_dash": True,
            "order_marker_regex": r"(?i)(?:\bOrder[_\s-]*(\d{4,})|#\s*(\d{4,}))",
        }

rules = load_rules_safe()
ORDER_RE = re.compile(rules.get("order_marker_regex", r"(?i)(?:\bOrder[_\s-]*(\d{4,})|#\s*(\d{4,}))"))

# ---------- утилиты ----------
def money_to_float(x: str) -> float:
    x = x.replace(" ", "").replace("\u00A0", "")
    return round(float(x.replace(",", ".")), 2)

def qty_to_int(x: str) -> int:
    return int(float(x.replace(" ", "").replace(",", ".")))

def extract_order(text: str) -> str | None:
    m = ORDER_RE.search(text)
    if m:
        return m.group(1) or m.group(2)
    return None

# ---------- извлечение текста (PyPDF2 -> OCR) ----------
def get_text_pages(pdf_bytes: bytes) -> list[str]:
    pages = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
    except Exception:
        reader = None

    for i in range(len(doc)):
        t = ""
        if reader:
            try:
                t = reader.pages[i].extract_text() or ""
            except Exception:
                t = ""
        if len(t.strip()) <= 50:
            # OCR
            pix = doc.load_page(i).get_pixmap(dpi=220)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            try:
                t = pytesseract.image_to_string(img, lang="eng+rus+lav")
            except Exception:
                t = pytesseract.image_to_string(img, lang="eng")
        pages.append(t)
    return pages

# ---------- парсер под твой формат ----------
def auto_parse(text_pages: list[str]) -> pd.DataFrame:
    """
    - MPN: последний 11-значный код или 81.XXXXX-YYYY
    - Qty: GAB … | … GAB | число прямо перед MPN (текущая/предыдущая строка)
    - Total: последняя НЕ нулевая денежная сумма слева от MPN (текущая/предыдущая строка), не равная qty
    - Order: контекст по #123456 / Order_123456
    - Дедуп по (Order, MPN)
    """
    # регулярки
    MPN_11 = re.compile(r"\b(\d{11})\b")
    MPN_DASH = re.compile(r"(?i)C?(\d{2}\.\d{5}-\d{3,5})")
    MONEY = re.compile(r"\d{1,3}(?:[\s\u00A0]?\d{3})*[.,]\d{2}")  # 2 знака после запятой, с пробелами в тысячах
    GAB_AFTER = re.compile(r"(?i)\bGAB\b[^0-9]{0,12}(\d+[\.,]?\d*)")   # GAB 48% 7,00
    GAB_BEFORE = re.compile(r"(?i)(\d+[\.,]?\d*)\s*\bGAB\b")           # 7,00 GAB

    # нормализованные строки
    lines = []
    for t in text_pages:
        for raw in t.splitlines():
            s = " ".join(raw.split())
            if s:
                lines.append(s)

    rows_by_key: dict[tuple[str, str], dict] = {}
    current_order = None

    for i, line in enumerate(lines):
        prev = lines[i - 1] if i > 0 else ""

        # обновляем order из текущей/предыдущей строки
        for src in (line, prev):
            m_ord = extract_order(src)
            if m_ord:
                current_order = m_ord

        # --- MPN: ищем ВСЕ кандидаты в строке, берём ПОСЛЕДНИЙ (правее всего) ---
        mpn_matches = [(m.group(1), m.span()) for m in MPN_11.finditer(line)]
        if not mpn_matches:
            mpn_matches = [(m.group(1), m.span()) for m in MPN_DASH.finditer(line)]
        if not mpn_matches:
            continue
        mpn_raw, mpn_span = mpn_matches[-1]
        mpn = cleanse_mpn(mpn_raw, rules)

        # --- QTY ---
        qty, qty_span = None, None

        # 1) по GAB в текущей строке
        m = GAB_AFTER.search(line) or GAB_BEFORE.search(line)
        if m:
            try:
                qty = qty_to_int(m.group(1)); qty_span = m.span(1)
            except Exception:
                qty = None

        # 2) число сразу перед MPN (текущая строка)
        if qty is None:
            before = line[:mpn_span[0]]
            m_pre = re.search(r"(\d{1,5})(?:[,\.]\d{1,2})?\s*$", before)
            if m_pre:
                try:
                    qty = qty_to_int(m_pre.group(1)); qty_span = (mpn_span[0]-len(m_pre.group(0))+m_pre.start(1), mpn_span[0]-len(m_pre.group(0))+m_pre.end(1))
                except Exception:
                    qty = None

        # 3) GAB/число-перед-MPN на предыдущей строке (если перенос)
        if qty is None and prev:
            m = GAB_AFTER.search(prev) or GAB_BEFORE.search(prev)
            if m:
                try:
                    qty = qty_to_int(m.group(1)); qty_span = (-1,-1)
                except Exception:
                    qty = None
            if qty is None:
                m_pre = re.search(r"(\d{1,5})(?:[,\.]\d{1,2})?\s*$", prev)
                if m_pre:
                    try:
                        qty = qty_to_int(m_pre.group(1)); qty_span = (-1,-1)
                    except Exception:
                        qty = None

        if qty is None:
            qty = 1

        # --- TOTAL ---
        def pick_total(src: str, only_left_of: int | None) -> tuple[float | None, tuple[int,int] | None]:
            best = None
            best_span = None
            for m in MONEY.finditer(src):
                span = m.span()
                if only_left_of is not None and span[0] >= only_left_of:
                    continue  # только суммы левее MPN
                val = m.group(0)
                # игнорируем сумму, совпадающую с qty (например 400,00 сразу перед MPN)
                if qty_span and span == qty_span:
                    continue
                try:
                    num = money_to_float(val)
                except Exception:
                    continue
                if num == 0:
                    continue
                # выбираем крайнюю правую
                if best is None or span[0] > best_span[0]:
                    best, best_span = num, span
            return best, best_span

        total, _ = pick_total(line, mpn_span[0])  # только слева от MPN
        if total is None and prev:
            total, _ = pick_total(prev, None)
        if total is None:
            total = 0.0

        key = (current_order or "", mpn)
        new_row = {
            "MPN": mpn,
            "Replacem": "",
            "Quantity": qty,
            "Totalsprice": total,
            "Order reference": current_order or "",
            "_line_idx": i,
        }

        # дедуп: оставляем лучшую запись
        if key in rows_by_key:
            old = rows_by_key[key]
            choose = False
            if old["Totalsprice"] == 0 and new_row["Totalsprice"] != 0:
                choose = True
            elif new_row["Totalsprice"] != 0 and old["Totalsprice"] != 0:
                # обе не ноль — оставляем ближайшую к MPN (из текущей строки предпочтительнее предыдущей)
                if new_row["_line_idx"] >= old["_line_idx"]:
                    choose = True
            elif new_row["Totalsprice"] == old["Totalsprice"]:
                if new_row["Quantity"] > old["Quantity"]:
                    choose = True
            if choose:
                rows_by_key[key] = new_row
        else:
            rows_by_key[key] = new_row

    if not rows_by_key:
        return pd.DataFrame(columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])

    df = pd.DataFrame(rows_by_key.values(), columns=["MPN","Replacem","Quantity","Totalsprice","Order reference","_line_idx"])
    df = df.drop(columns=["_line_idx"], errors="ignore")
    # сортировка для удобства
    with pd.option_context("mode.copy_on_write", True):
        df["Order reference"] = df["Order reference"].astype(str)
    df = df.sort_values(["Order reference","MPN"]).reset_index(drop=True)
    return df

# ---------- UI ----------
pdf_file = st.file_uploader("Загрузить счёт (PDF)", type=["pdf"])
tpl_file = st.file_uploader("Шаблон Excel (необязательно)", type=["xlsx"])

if not pdf_file:
    st.info("Загрузите PDF — таблица сформируется автоматически.")
else:
    pages_text = get_text_pages(pdf_file.read())
    df = auto_parse(pages_text)

    with st.expander("DEBUG (распознанный текст, первые 3000 символов)", expanded=False):
        st.text("\n\n".join(pages_text)[:3000])

    st.subheader("Предпросмотр")
    st.dataframe(df, use_container_width=True)

    if st.button("Скачать waybill"):
        if tpl_file:
            wb = load_workbook(tpl_file); ws = wb.active
        else:
            wb = Workbook(); ws = wb.active
            ws["A1"]="MPN"; ws["B1"]="Replacem"; ws["C1"]="Quantity"; ws["D1"]="Totalsprice"; ws["E1"]="Order reference"

        # очистка и запись
        for r in range(2, 3000):
            for c in range(1, 6):
                ws.cell(row=r, column=c, value=None)

        for i, row in enumerate(df.values.tolist(), start=2):
            for j, v in enumerate(row, start=1):
                ws.cell(row=i, column=j, value=v)

        bio = io.BytesIO(); wb.save(bio)
        st.download_button(
            "Скачать waybill.xlsx",
            data=bio.getvalue(),
            file_name="waybill.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
