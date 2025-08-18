import streamlit as st, re, io, pandas as pd
from PyPDF2 import PdfReader
import fitz  # PyMuPDF для OCR
import pytesseract
from PIL import Image
from openpyxl import load_workbook, Workbook
from utils import load_profiles, cleanse_mpn

st.set_page_config(page_title="Waybill Maker", page_icon="📦", layout="wide")
st.title("📦 Waybill Maker")

# ---------- Правила из YAML с фоллбэком ----------
def load_rules_safe():
    try:
        profiles = load_profiles("supplier_profiles.yaml")
        return profiles.get("default", {})
    except Exception:
        return {
            "remove_leading_C_in_mpn": True,
            "materom_mpn_before_dash": True,
            "order_marker_regex": r"(?i)(?:\bOrder[_\s-]*(\d{4,})|#\s*(\d{4,}))",
            # MPN: 81.XXXXX-YYYY ИЛИ 11–12 цифр
            "mpn_patterns": [
                r"(?i)C?(\d{2}\.\d{5}-\d{3,5})",
                r"\b(\d{11,12})\b",
            ],
            # Qty: GAB 7,00 / 7,00 GAB / число перед MPN
            "qty_patterns": [
                r"(?i)\bGAB\b[^0-9]{0,12}(\d+[\.,]?\d*)",
                r"(?i)(\d+[\.,]?\d*)\s*\bGAB\b",
                r"(\d{1,5})(?:[,\.]00)?(?=\s*\d{6,}\b)",
            ],
            # Итог: последняя НЕ нулевая денежная сумма (учёт пробелов в тысячах)
            "total_patterns": [
                r"(?i)(\d{1,3}(?:[\s\u00A0]?\d{3})*[.,]\d{2})(?!.*\d{1,3}(?:[\s\u00A0]?\d{3})*[.,]\d{2})",
            ],
        }

rules = load_rules_safe()
order_re = re.compile(
    rules.get("order_marker_regex", r"(?i)(?:\bOrder[_\s-]*(\d{4,})|#\s*(\d{4,}))")
)

# ---------- Утилиты ----------
def money_to_float(x: str) -> float:
    x = x.replace(" ", "").replace("\u00A0", "")
    return round(float(x.replace(",", ".")), 2)

def qty_to_int(x: str) -> int:
    return int(float(x.replace(" ", "").replace(",", ".")))

def extract_order(line: str) -> str | None:
    m = order_re.search(line)
    if m:
        return m.group(1) or m.group(2)
    return None

# ---------- Извлечение текста (текст → OCR fallback) ----------
def get_all_text(pdf_bytes: bytes) -> list[str]:
    texts = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
    except Exception:
        reader = None

    for i in range(len(doc)):
        text = ""
        if reader:
            try:
                text = reader.pages[i].extract_text() or ""
            except Exception:
                text = ""
        if len(text.strip()) <= 50:
            pix = doc.load_page(i).get_pixmap(dpi=220)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            try:
                text = pytesseract.image_to_string(img, lang="eng+rus+lav")
            except Exception:
                text = pytesseract.image_to_string(img, lang="eng")
        texts.append(text)
    return texts

# ---------- Парсер (заточен под твой формат) ----------
def auto_parse(text_pages: list[str]) -> pd.DataFrame:
    """
    - MPN: последний 11-значный код (или 81.XXXXX-YYYY)
    - Qty: из GAB … / … GAB или число перед MPN (иначе 1)
    - Total: последняя НЕ нулевая сумма в строке (fallback: предыдущая строка)
    - Order: контекст по #123456 / Order_123456
    - Дедуп по (Order, MPN)
    """
    order_re_local = re.compile(rules.get("order_marker_regex", r"(?i)(?:\bOrder[_\s-]*(\d{4,})|#\s*(\d{4,}))"))
    money_re = re.compile(r"\d{1,3}(?:[\s\u00A0]?\d{3})*[.,]\d{2}")
    mpn11_re = re.compile(r"\b(\d{11})\b")
    mpn_dash_re = re.compile(r"(?i)C?(\d{2}\.\d{5}-\d{3,5})")
    gab_after_re = re.compile(r"(?i)\bGAB\b[^0-9]{0,12}(\d+[\.,]?\d*)")
    gab_before_re = re.compile(r"(?i)(\d+[\.,]?\d*)\s*\bGAB\b")

    # плоский список строк
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

        # Order из текущей/предыдущей строки
        for src in (line, prev):
            m_ord = order_re_local.search(src)
            if m_ord:
                current_order = (m_ord.group(1) or m_ord.group(2))

        # MPN: предпочитаем 11 цифр; если нет — формат с дефисом
        mpn_candidates = [m.group(1) for m in mpn11_re.finditer(line)]
        if not mpn_candidates:
            mpn_candidates = [m.group(1) for m in mpn_dash_re.finditer(line)]
        if not mpn_candidates:
            continue
        mpn = cleanse_mpn(mpn_candidates[-1], rules)

        # Quantity
        qty = None
        m = gab_after_re.search(line) or gab_before_re.search(line)
        if m:
            try:
                qty = qty_to_int(m.group(1))
            except Exception:
                qty = None

        if qty is None:
            m_pre = re.search(r"(\d{1,5})(?:[,\.]\d{1,2})?\s+" + re.escape(mpn) + r"\b", line)
            if m_pre:
                try:
                    qty = qty_to_int(m_pre.group(1))
                except Exception:
                    qty = None

        if qty is None and prev:
            m = gab_after_re.search(prev) or gab_before_re.search(prev)
            if m:
                try:
                    qty = qty_to_int(m.group(1))
                except Exception:
                    qty = None
            if qty is None:
                m_pre = re.search(r"(\d{1,5})(?:[,\.]\d{1,2})?\s+" + re.escape(mpn) + r"\b", prev)
                if m_pre:
                    try:
                        qty = qty_to_int(m_pre.group(1))
                    except Exception:
                        qty = None

        if qty is None:
            qty = 1

        # Totalsprice: последняя НЕ нулевая в строке, иначе в предыдущей
        total = None
        monies = money_re.findall(line)
        if monies:
            for mny in reversed(monies):
                if mny not in ("0,00", "0.00"):
                    try:
                        total = money_to_float(mny)
                        break
                    except Exception:
                        pass
        if total is None and prev:
            monies = money_re.findall(prev)
            for mny in reversed(monies):
                if mny not in ("0,00", "0.00"):
                    try:
                        total = money_to_float(mny)
                        break
                    except Exception:
                        pass
        if total is None:
            total = 0.0

        key = (current_order or "", mpn)
        new_row = {
            "MPN": mpn,
            "Replacem": "",
            "Quantity": qty,
            "Totalsprice": total,
            "Order reference": current_order or "",
        }

        # дедуп: оставляем запись с ненулевой ценой / большим qty / либо последнюю
        if key in rows_by_key:
            old = rows_by_key[key]
            choose = False
            if old["Totalsprice"] == 0 and new_row["Totalsprice"] != 0:
                choose = True
            elif new_row["Totalsprice"] == old["Totalsprice"]:
                if new_row["Quantity"] > old["Quantity"]:
                    choose = True
            elif new_row["Totalsprice"] != 0 and old["Totalsprice"] != 0:
                choose = True  # если обе не ноль, берём последнюю
            if choose:
                rows_by_key[key] = new_row
        else:
            rows_by_key[key] = new_row

    if not rows_by_key:
        return pd.DataFrame(columns=["MPN", "Replacem", "Quantity", "Totalsprice", "Order reference"])

    df = pd.DataFrame(rows_by_key.values(), columns=["MPN", "Replacem", "Quantity", "Totalsprice", "Order reference"])
    # сортировка для удобства
    with pd.option_context("mode.copy_on_write", True):
        df["Order reference"] = df["Order reference"].astype(str)
    df = df.sort_values(["Order reference", "MPN"]).reset_index(drop=True)
    return df

# ---------- UI ----------
pdf_file = st.file_uploader("Загрузить счёт (PDF)", type=["pdf"])
tpl_file = st.file_uploader("Шаблон Excel (необязательно)", type=["xlsx"])

if not pdf_file:
    st.info("Загрузите PDF-счёт. Я сам распознаю и соберу Excel по правилам.")
else:
    pages_text = get_all_text(pdf_file.read())
    df = auto_parse(pages_text)

    with st.expander("DEBUG (распознанный текст, первые 3000 символов)", expanded=False):
        st.text("\n\n".join(pages_text)[:3000])

    st.subheader("Предпросмотр (авто)")
    st.dataframe(df, use_container_width=True)

    if st.button("Скачать waybill"):
        if tpl_file:
            wb = load_workbook(tpl_file); ws = wb.active
        else:
            wb = Workbook(); ws = wb.active
            ws["A1"] = "MPN"; ws["B1"] = "Replacem"; ws["C1"] = "Quantity"
            ws["D1"] = "Totalsprice"; ws["E1"] = "Order reference"

        # очистка и запись
        for r in range(2, 3000):
            for c in range(1, 6):
                ws.cell(row=r, column=c, value=None)

        for i, row in enumerate(df.values.tolist(), start=2):
            for j, value in enumerate(row, start=1):
                ws.cell(row=i, column=j, value=value)

        bio = io.BytesIO(); wb.save(bio)
        st.download_button(
            "Скачать waybill.xlsx",
            data=bio.getvalue(),
            file_name="waybill.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
