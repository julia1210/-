"""
컨기부 창고 품목의 단가 매칭 여부 확인
"""
import sys
from datetime import datetime
from collections import defaultdict
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from ecount_inventory import get_inventory_by_location, extract_rows
from ecount_items import fetch_item_master
from ecount_warehouses import WH_GROUPS

base_date = datetime.now().strftime("%Y%m%d")
kongi_codes = set(WH_GROUPS["컨기부"])

print("품목 마스터 로드 중...")
item_master, barcode_index = fetch_item_master()

print("전체 재고 조회 중...")
raw = get_inventory_by_location(base_date)
rows = extract_rows(raw)

print(f"\n=== 컨기부 창고({kongi_codes}) 품목 단가 확인 ===")
print(f"{'품목코드':<20} {'입고단가':>12}  {'창고':>6}  품목명")
print("-" * 80)

found = 0
no_price = 0
for row in rows:
    if not isinstance(row, dict):
        continue
    wh_cd = str(row.get("WH_CD") or "").strip()
    if wh_cd not in kongi_codes:
        continue
    prod_cd = str(row.get("PROD_CD") or "").strip()
    bal_qty = row.get("BAL_QTY", 0)
    master = item_master.get(prod_cd, {})
    in_price = master.get("in_price", None)
    prod_des = master.get("prod_des") or str(row.get("PROD_DES") or "")
    found += 1
    if not in_price:
        no_price += 1
        print(f"  {prod_cd:<20} {'[단가없음]':>12}  {wh_cd:>6}  {prod_des[:40]}")

print(f"\n총 {found}건 중 단가 없음: {no_price}건")

print("\n=== 품목 마스터에 없는 품목코드 ===")
missing = 0
for row in rows:
    if not isinstance(row, dict):
        continue
    wh_cd = str(row.get("WH_CD") or "").strip()
    if wh_cd not in kongi_codes:
        continue
    prod_cd = str(row.get("PROD_CD") or "").strip()
    if prod_cd not in item_master:
        print(f"  {prod_cd}  (재고API에는 있지만 품목마스터에 없음)")
        missing += 1
if missing == 0:
    print("  없음 (모든 품목코드가 품목 마스터에 존재)")
