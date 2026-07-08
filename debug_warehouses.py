"""
전체 재고 조회 시 실제로 어떤 창고코드(WH_CD)가 오는지 확인
"""
import sys
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from datetime import datetime
from collections import defaultdict
from ecount_inventory import get_inventory_by_location, extract_rows
from ecount_warehouses import WH_GROUPS

base_date = datetime.now().strftime("%Y%m%d")

print("전체 창고 재고 조회 중...")
raw = get_inventory_by_location(base_date)
rows = extract_rows(raw)
print(f"총 {len(rows)}건\n")

# WH_CD별 품목 수 집계
wh_counts = defaultdict(int)
for row in rows:
    if isinstance(row, dict):
        wh_cd = str(row.get("WH_CD") or "").strip()
        wh_counts[wh_cd] += 1

print("=== 응답에 포함된 WH_CD 목록 ===")
for wh_cd, cnt in sorted(wh_counts.items()):
    print(f"  WH_CD='{wh_cd}' → {cnt}개 품목")

print("\n=== 설정된 창고코드와 매칭 여부 ===")
all_configured = {code for codes in WH_GROUPS.values() for code in codes}
found = set(wh_counts.keys())

for code in sorted(all_configured):
    status = "OK" if code in found else "NOT FOUND"
    print(f"  {code} → {status}")

print("\n=== 응답에는 있지만 설정에 없는 WH_CD ===")
extra = found - all_configured - {""}
for code in sorted(extra):
    print(f"  {code} ({wh_counts[code]}개)")
