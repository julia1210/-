"""
CCS 브랜드 디버그 스크립트
- 재고 API에서 CCS 관련 품목 찾아 브랜드 분류 과정 추적
"""
import sys, json, unicodedata, re
sys.path.insert(0, '.')

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from ecount_inventory import get_inventory_by_location, extract_rows
from ecount_items import fetch_item_master
from ecount_warehouses import WH_GROUPS, BRAND_SHEET_MAP

def _nfc(s):
    return unicodedata.normalize("NFC", str(s).strip())

NAMED_BRANDS = {
    "굿스마일", "메가하우스", "부시로드", "블로키", "반다이남코",
    "핫토이", "하비재팬", "CCS TOYS", "CCS", "P001",
    "코코파", "코코파스튜디오", "마그넷",
}
KNOWN_BRANDS = list(BRAND_SHEET_MAP.keys())

def extract_brand(prod_des):
    if not prod_des:
        return "기타브랜드"
    cleaned = _nfc(prod_des.strip())
    m = re.match(r"\[([^\]]+)\]", cleaned)
    if m:
        raw = _nfc(m.group(1).strip())
        for brand in KNOWN_BRANDS:
            if raw == _nfc(brand) or raw.lower() == _nfc(brand).lower():
                return brand
        return raw
    cl = cleaned.lower()
    for brand in sorted(KNOWN_BRANDS, key=len, reverse=True):
        b_nfc = _nfc(brand).lower()
        if cl.startswith(b_nfc + " ") or cl.startswith(b_nfc + "_") or cl == b_nfc:
            return brand
    return "기타브랜드"

from datetime import datetime
base_date = datetime.now().strftime("%Y%m%d")

print("품목 마스터 로드...")
item_master, barcode_index = fetch_item_master()

print("재고 조회...")
from ecount_inventory import get_inventory_by_location, extract_rows
raw = get_inventory_by_location(base_date)
all_rows = extract_rows(raw)
print(f"총 {len(all_rows)}건")

wh_to_group = {}
for group, codes in WH_GROUPS.items():
    for code in codes:
        wh_to_group[code] = group

# CCS 관련 행 찾기
print("\n=== CCS 관련 재고 행 (PROD_DES에 'CCS' 포함) ===")
ccs_rows = [r for r in all_rows if isinstance(r, dict) and "ccs" in str(r.get("PROD_DES","")).lower()]
print(f"총 {len(ccs_rows)}개 행 발견")
for r in ccs_rows[:20]:
    wh = str(r.get("WH_CD",""))
    group = wh_to_group.get(wh, "(그룹없음)")
    prod_cd = str(r.get("PROD_CD",""))
    prod_des = str(r.get("PROD_DES",""))
    bal = r.get("BAL_QTY","")
    brand_extracted = extract_brand(prod_des)
    master = item_master.get(prod_cd) or barcode_index.get(prod_cd, {})
    master_brand = master.get("brand","(없음)")
    print(f"  WH={wh}({group}) PROD_CD={prod_cd} BAL={bal}")
    print(f"    PROD_DES={repr(prod_des)}")
    print(f"    첫글자 ord={[ord(c) for c in prod_des[:5]]}")
    print(f"    extract_brand -> {repr(brand_extracted)}")
    print(f"    item_master brand -> {repr(master_brand)}")
    print(f"    BRAND_SHEET_MAP.get(extract) -> {repr(BRAND_SHEET_MAP.get(brand_extracted))}")
    print()

if not ccs_rows:
    print("  CCS 관련 품목 없음!")
    print("\n  첫 5개 재고 행:")
    for r in all_rows[:5]:
        print(" ", r)
