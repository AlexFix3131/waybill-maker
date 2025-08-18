def parse_pdf(pdf_bytes: bytes) -> pd.DataFrame:
    rows, current_order_digits = [], None
    rules_local = rules
    order_re_local = order_re

    # --- собираем текст ПОСТРАНИЧНО: если мало текста — OCR для этой страницы
    page_texts = []
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        num_pages = len(reader.pages)
    except Exception:
        reader = None
        num_pages = 0

    def page_to_text_with_ocr(page_index: int) -> str:
        # пробуем извлечь как текст
        if reader:
            try:
                t = reader.pages[page_index].extract_text() or ""
                if len(t.strip()) > 50:
                    return t
            except Exception:
                pass
        # если текста мало/нет — OCR
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page = doc.load_page(page_index)
        pix = page.get_pixmap(dpi=220)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        try:
            return pytesseract.image_to_string(img, lang="eng+rus+lav")
        except Exception:
            return pytesseract.image_to_string(img, lang="eng")

    if num_pages > 0:
        for i in range(num_pages):
            page_texts.append(page_to_text_with_ocr(i))
    else:
        # fallback: OCR всего файла одной картинкой
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for p in doc:
            pix = p.get_pixmap(dpi=220)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            try:
                page_texts.append(pytesseract.image_to_string(img, lang="eng+rus+lav"))
            except Exception:
                page_texts.append(pytesseract.image_to_string(img, lang="eng"))

    # --- DEBUG: показываем, что реально распознали (обрезаем для UI)
    st.text_area("DEBUG: что удалось вытащить из PDF/OCR",
                 "\n\n--- PAGE SPLIT ---\n\n".join([t[:2000] for t in page_texts]),
                 height=260)

    # вспомогалки
    def find_first(pattern_key: str, line: str, conv=None):
        for patt in rules_local.get(pattern_key, []):
            m = re.search(patt, line)
            if m:
                val = m.group(1)
                if conv:
                    try:
                        return conv(val)
                    except Exception:
                        return None
                return val
        return None

    money_any_re = re.compile(r"\d{1,3}(?:[\s\u00A0]?\d{3})*[.,]\d{2}")

    # --- разбор построчно
    for text in page_texts:
        for raw_line in text.splitlines():
            line = " ".join(raw_line.split())

            # Order: "#123456" ИЛИ "Order_123456"
            m_order = order_re_local.search(line)
            if m_order:
                current_order_digits = (m_order.group(1) or m_order.group(2))

            # MPN
            mpn = find_first("mpn_patterns", line)
            if not mpn:
                continue
            mpn = cleanse_mpn(mpn, rules_local)

            # Quantity
            def to_int(x): return int(float(x.replace(" ", "").replace(",", ".")))
            qty = find_first("qty_patterns", line, to_int)
            if qty is None:
                m_pre = re.search(r"(\d{1,5})(?:[,\.]00)?\s*"+re.escape(mpn)+r"\s*$", line)
                if m_pre:
                    try: qty = int(m_pre.group(1))
                    except: qty = None
            if qty is None:
                qty = 1

            # Totalsprice: последняя ненулевая сумма (учитываем пробелы в тысячах)
            def to_money(x):
                x = x.replace(" ", "").replace("\u00A0", "")
                return round(float(x.replace(",", ".")), 2)

            total = find_first("total_patterns", line, to_money)
            if total is None:
                all_money = money_any_re.findall(line)
                if all_money:
                    last = all_money[-1]
                    if last not in ("0,00", "0.00"):
                        try: total = to_money(last)
                        except: total = None
            if total is None:
                total = 0.0

            rows.append([mpn, "", qty, total, current_order_digits or ""])

    return pd.DataFrame(rows, columns=["MPN", "Replacem", "Quantity", "Totalsprice", "Order reference"])
