"""
이카운트 품목 마스터의 CONT1(브랜드) 값 현황 확인
"""
import sys, json
from collections import Counter
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from ecount_items import fetch_item_master

print("품목 마스터 로드 중...")
items = fetch_item_master()
print(f"총 {len(items)}개 품목\n")

brand_count = Counter()
for info in items.values():
    b = info.get("brand") or "(없음)"
    brand_count[b] += 1

print("=== CONT1(브랜드) 값별 품목 수 ===")
for brand, cnt in brand_count.most_common():
    print(f"  '{brand}': {cnt}개")
