
import re, yaml

def load_profiles(path="supplier_profiles.yaml"):
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def only_digits_from_order(txt):
    m = re.search(r"(?i)\bOrder[_\s-]*(\d{4,})", txt or "")
    return m.group(1) if m else None

def cleanse_mpn(mpn, rules):
    if not mpn: return mpn
    mpn = mpn.strip()
    if rules.get("remove_leading_C_in_mpn", True) and mpn.startswith(("C","c")):
        mpn = mpn[1:]
    if rules.get("materom_mpn_before_dash", False):
        m = re.match(r"^(\d+)-", mpn)
        if m: return m.group(1)
    return mpn

def ital_express_fix(ref, rules):
    if not ref: return ref
    if rules.get("ital_express_strip_prefix", False):
        m = re.match(r"^\d+\/(\w+)$", ref)
        if m: return m.group(1)
    return ref
