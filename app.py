import streamlit as st, re, io, pandas as pd
from PyPDF2 import PdfReader
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from openpyxl import load_workbook, Workbook
from utils import load_profiles, cleanse_mpn

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

# ========== правила (yaml -> fallback) ==========
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

RULES = load_rules_safe()
ORDER_RE = re.compile(RULES.get("order_marker_regex", r"(?i)(?:\bOrder[_\s-]*(\d{4,})|#\s*(\d{4,}))"))

# ========== утилиты ==========
def money_to_float(x: str) -> float:
    x = x.replace(" ", "").replace("\u00A0", "")
    return round(float(x.replace(",", ".")), 2)

def qty_to_int(x: str) -> int:
    return int(float(x.replace(" ", "").replace(",", ".")))

def extract_order(s: str) -> str | None:
    m = ORDER_RE.search(s)
    if m:
        return m.group(1) or m.group(2)
    return None

# ========== извлечение текста (PyPDF2 -> OCR) ==========
def get_text_pages(pdf_bytes: bytes) -> list[str]:
    out = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
    except Exception:
        reader = None

    for i in range(len(doc)):
        txt = ""
        if reader:
            try:
                txt = reader.pages[i].extract_text() or ""
            except Exception:
                txt = ""
        if len(txt.strip()) <= 50:
            pix = doc.load_page(i).get_pixmap(dpi=220)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            try:
                txt = pytesseract.image_to_string(img, lang="eng+rus+lav")
            except Exception:
                txt = pytesseract.image_to_string(img, lang="eng")
        out.append(txt)
    return out

# ========== ПАРСЕР: якорь = строка с MPN ==========
def auto_parse(text_pages: list[str]) -> pd.DataFrame:
    """
    Алгоритм:
      1) строим список строк;
      2) проход слева-направо: держим текущий Order;
      3) когда встречаем строку с MPN (11 цифр ИЛИ 81.XXXXX-YYYY) — это якорь;
      4) в окне [i-3 .. i] ищем Quantity (GAB или число прямо перед MPN) и Totalsprice
         (последняя НЕ нулевая сумма левее MPN; если нет — просто последняя не нулевая в окне);
      5) дедуп по (Order, MPN): оставляем с ненулевой ценой/бОльшим qty/ближе к MPN.
    """
    # regex
    MPN_11   = re.compile(r"\b(\d{11})\b")
    MPN_DASH = re.compile(r"(?i)C?(\d{2}\.\d{5}-\d{3,5})")
    MONEY    = re.compile(r"\d{1,3}(?:[\s\u00A0]?\d{3})*[.,]\d{2}")
    GAB_A    = re.compile(r"(?i)\bGAB\b[^0-9]{0,12}(\d+[\.,]?\d*)")  # GAB … 7,00
    GAB_B    = re.compile(r"(?i)(\d+[\.,]?\d*)\s*\bGAB\b")          # 7,00 GAB

    # шум (банки/адреса) — чтобы не ловить мусор
    NOISE = re.compile(r"(?i)\b(IBAN|SWIFT|bank|banka|konto|account|address|adrese|PVN|VAT|invoice|rekins|rekīns|tel\.?|email)\b")

    # плоские строки
    lines = []
    for t in text_pages:
        for raw in t.splitlines():
            s = " ".join(raw.split())
            if s:
                lines.append(s)

    # индексация заказов для скорости
    orders_at = {}
    for i, s in enumerate(lines):
        m = extract_order(s)
        if m:
            orders_at[i] = m

    def order_for_index(idx: int) -> str | None:
        # ближайший order слева
        for j in range(idx, -1, -1):
            if j in orders_at:
                return orders_at[j]
        return None

    def find_mpn_in_line(s: str):
        m = list(MPN_11.finditer(s))
        if m:
            mm = m[-1]  # самый правый
            return cleanse_mpn(mm.group(1), RULES), mm.span()
        m = list(MPN_DASH.finditer(s))
        if m:
            mm = m[-1]
            return cleanse_mpn(mm.group(1), RULES), mm.span()
        return None, None

    def find_qty(window_lines: list[str], anchor_line: str, mpn_span):
        # 1) GAB в окне (с конца)
        for s in reversed(window_lines):
            m = GAB_A.search(s) or GAB_B.search(s)
            if m:
                try:
                    return qty_to_int(m.group(1))
                except Exception:
                    pass
        # 2) число прямо перед MPN в якорной строке
        if mpn_span:
            left = anchor_line[:mpn_span[0]]
            mpre = re.search(r"(\d{1,5})(?:[,\.]\d{1,2})?\s*$", left)
            if mpre:
                try:
                    return qty_to_int(mpre.group(1))
                except Exception:
                    pass
        # 3) число в конце строки в окне (часто «… 400,00»)
        for s in reversed(window_lines):
            mpre = re.search(r"(\d{1,5})(?:[,\.]\d{1,2})?\s*$", s)
            if mpre:
                try:
                    return qty_to_int(mpre.group(1))
                except Exception:
                    pass
        return 1

    def find_total(window_lines: list[str], anchor_line: str, mpn_span, qty_val: int):
        # пытаемся взять сумму ЛЕВЕЕ MPN в якорной строке — это самый надёжный вариант
        if mpn_span:
            left = anchor_line[:mpn_span[0]]
            monies = list(MONEY.finditer(left))
            for m in reversed(monies):  # самая правая слева от MPN
                val = m.group(0)
                try:
                    num = money_to_float(val)
                except Exception:
                    continue
                if num == 0:
                    continue
                # если число равно qty (типа 400,00 прямо перед MPN) — пропускаем
                if abs(num - qty_val) < 1e-9:
                    continue
                return num
        # иначе берём последнюю НЕ нулевую сумму в окне
        for s in reversed(window_lines):
            monies = list(MONEY.finditer(s))
            for m in reversed(monies):
                try:
                    num = money_to_float(m.group(0))
                except Exception:
                    continue
                if num != 0 and abs(num - qty_val) >= 1e-9:
                    return num
        return 0.0

    rows_by_key: dict[tuple[str, str], dict] = {}

    for i, line in enumerate(lines):
        if NOISE.search(line):
            continue

        mpn, span = find_mpn_in_line(line)
        if not mpn:
            continue

        # окно для контекста: три строки выше и текущая
        win_start = max(0, i - 3)
        window = lines[win_start:i + 1]

        qty = find_qty(window, line, span)
        order = order_for_index(i)
        total = find_total(window, line, span, qty)

        rec = {
            "MPN": mpn,
            "Replacem": "",
            "Quantity": qty,
            "Totalsprice": total,
            "Order reference": order or "",
            "_i": i,
        }
        key = (rec["Order reference"], mpn)

        # дедуп: оставляем лучшую запись
        if key in rows_by_key:
            old = rows_by_key[key]
            choose = False
            if old["Totalsprice"] == 0 and rec["Totalsprice"] != 0:
                choose = True
            elif rec["Totalsprice"] != 0 and old["Totalsprice"] != 0:
                # обе не ноль — берём ту, что ближе (позже)
                if rec["_i"] >= old["_i"]:
                    choose = True
            elif rec["Totalsprice"] == old["Totalsprice"]:
                if rec["Quantity"] > old["Quantity"]:
                    choose = True
            if choose:
                rows_by_key[key] = rec
        else:
            rows_by_key[key] = rec

    if not rows_by_key:
        return pd.DataFrame(columns=["MPN","Replacem","Quantity","Totalsprice","Order reference"])

    df = pd.DataFrame(rows_by_key.values(), columns=["MPN","Replacem","Quantity","Totalsprice","Order reference","_i"])
    df = df.drop(columns=["_i"], errors="ignore")
    with pd.option_context("mode.copy_on_write", True):
        df["Order reference"] = df["Order reference"].astype(str)
    df = df.sort_values(["Order reference", "MPN"]).reset_index(drop=True)
    return df

# ========== UI ==========
pdf_file = st.file_uploader("Загрузить счёт (PDF)", type=["pdf"])
tpl_file = st.file_uploader("Шаблон Excel (необязательно)", type=["xlsx"])

if not pdf_file:
    st.info("Загрузите PDF — нужные поля будут собраны в таблицу автоматически.")
else:
    pages_text = get_text_pages(pdf_file.read())
    df = auto_parse(pages_text)

    with st.expander("DEBUG (распознанный текст, первые 3000 символов)"):
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
