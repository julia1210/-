"""
이카운트 재고현황 조회 실행 스크립트
─────────────────────────────────────────
창고별 재고현황 API 를 호출하고, 결과 요약을 보여준 뒤
원본 응답 전체를 output 폴더에 JSON 으로 저장한다.

사용 예:
  python query_ecount_inventory.py                # 오늘 날짜, 전체 창고
  python query_ecount_inventory.py --wh 20002     # 20002(디아이 판매) 창고만
  python query_ecount_inventory.py --date 20260616 --wh 20002
"""

import sys
import json
import argparse
import traceback
from pathlib import Path
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from ecount_inventory import get_inventory_by_location, extract_rows
from ecount_config import load_ecount_config


def main():
    parser = argparse.ArgumentParser(description="이카운트 재고현황 조회")
    parser.add_argument("--date", default=datetime.now().strftime("%Y%m%d"),
                        help="기준일자 YYYYMMDD (기본: 오늘)")
    parser.add_argument("--wh", default=None, help="창고코드 (예: 20002). 생략 시 전체")
    parser.add_argument("--prod", default=None, help="품목코드. 생략 시 전체")
    args = parser.parse_args()

    server_label = "테스트(sboapi)" if load_ecount_config().is_test else "실서버(oapi)"
    print("=" * 56)
    print(f" 이카운트 재고현황 조회 [{server_label}]")
    print("=" * 56)
    print(f"  기준일자 : {args.date}")
    print(f"  창고코드 : {args.wh or '(전체)'}")
    print(f"  품목코드 : {args.prod or '(전체)'}")

    try:
        raw = get_inventory_by_location(args.date, wh_cd=args.wh, prod_cd=args.prod)
    except Exception as e:
        print(f"\n[실패] API 호출 중 오류")
        print(f"  → {e}")
        return 1

    status = raw.get("Status")
    error = raw.get("Error")
    errors = raw.get("Errors")
    print(f"\n[응답 상태] Status={status}")
    if error:
        print(f"  Error: {json.dumps(error, ensure_ascii=False)}")
    if errors:
        print(f"  Errors: {json.dumps(errors, ensure_ascii=False)}")

    rows = extract_rows(raw)
    print(f"\n[결과] 재고 행 수: {len(rows)} 건")

    for i, row in enumerate(rows[:5], 1):
        if isinstance(row, dict):
            preview = {k: row[k] for k in list(row.keys())[:8]}
            print(f"  {i}. {json.dumps(preview, ensure_ascii=False)}")
        else:
            print(f"  {i}. {row}")

    out_dir = Path("output")
    out_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = out_dir / f"ecount_inventory_raw_{stamp}.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(raw, f, ensure_ascii=False, indent=2)
    print(f"\n[저장] 원본 응답 전체 → {out_path}")

    print("\n" + "=" * 56)
    if status in ("200", 200) and not error:
        print(" ✅ 호출 성공")
    else:
        print(" ⚠️ 호출은 됐으나 응답에 상태/오류가 있습니다. 위 내용을 확인하세요.")
    print("=" * 56)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        print("\n[예상치 못한 오류]")
        traceback.print_exc()
        sys.exit(1)
