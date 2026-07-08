"""
이카운트 품목조회 API 테스트
─────────────────────────────────────────
기초등록API > 품목조회 를 호출해서
PROD_CD / 단가 / 바코드 등 어떤 필드가 오는지 확인한다.
"""

import sys
import json

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import requests
from ecount_auth import get_session

TIMEOUT = 30

# 이카운트 품목조회 API 경로 (기초등록API)
ITEM_LIST_PATH = "/OAPI/V2/InventoryBasic/GetBasicProductsList"


def main():
    print("=" * 56)
    print(" 이카운트 품목조회 API 테스트")
    print("=" * 56)

    sess = get_session(force=True)
    url = f"{sess['base_url']}{ITEM_LIST_PATH}?SESSION_ID={sess['session_id']}"
    print(f"\nURL: {url}")

    # 일단 빈 body로 호출 (전체 품목)
    body = {}
    print(f"body: {json.dumps(body, ensure_ascii=False)}")

    resp = requests.post(url, json=body, timeout=TIMEOUT)
    print(f"\nHTTP 상태코드: {resp.status_code}")

    try:
        data = resp.json()
    except Exception:
        print("JSON 파싱 실패. 원본 응답:")
        print(resp.text[:500])
        return 1

    status = data.get("Status")
    error = data.get("Error")
    print(f"Status: {status}")
    if error:
        print(f"Error: {json.dumps(error, ensure_ascii=False)}")

    # 응답 구조 파악
    raw_data = data.get("Data")
    print(f"\n응답 Data 타입: {type(raw_data)}")

    rows = []
    if isinstance(raw_data, list):
        rows = raw_data
    elif isinstance(raw_data, dict):
        for key in ("Result", "Datas", "Data", "ResultData", "List"):
            val = raw_data.get(key)
            if isinstance(val, list):
                rows = val
                print(f"  → Data.{key} 에서 리스트 발견")
                break

    print(f"\n총 {len(rows)}개 품목")

    if rows:
        print("\n[첫 번째 품목 전체 필드]")
        first = rows[0]
        if isinstance(first, dict):
            for k, v in first.items():
                print(f"  {k}: {v}")
        print("\n[처음 3개 품목 요약]")
        for i, row in enumerate(rows[:3], 1):
            if isinstance(row, dict):
                # 주요 필드만 출력
                keys = ["PROD_CD", "PROD_DES", "SIZE_DES", "IN_PRICE", "BAR_CODE",
                        "OUT_PRICE", "PROD_TYPE", "UNIT"]
                summary = {k: row.get(k) for k in keys if row.get(k) is not None}
                print(f"  {i}. {json.dumps(summary, ensure_ascii=False)}")

    # 전체 응답 저장
    import os
    os.makedirs("output", exist_ok=True)
    out_path = "output/item_api_raw.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"\n전체 응답 저장: {out_path}")

    print("\n" + "=" * 56)
    if status in ("200", 200) and not error:
        print(" OK - 위 필드 목록에서 단가/바코드 필드 확인하세요")
    else:
        print(" 호출 실패 또는 오류 - 위 내용 확인하세요")
    print("=" * 56)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        print(f"\n오류: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
