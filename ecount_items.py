"""
이카운트 품목 마스터 조회 및 캐싱
─────────────────────────────────────────
GetBasicProductsList API로 전체 품목을 받아
PROD_CD → {prod_des, brand, in_price, bar_code} 딕셔너리를 반환한다.

결과는 .ecount_items_cache.json 에 캐싱해 매번 호출하지 않는다.
(품목 마스터는 재고보다 변동이 적으므로 하루 1회 갱신으로 충분)
"""

import json
import time
import requests
from pathlib import Path

from ecount_auth import get_session

TIMEOUT = 60
ITEM_CACHE_FILE = Path(".ecount_items_cache.json")
ITEM_CACHE_TTL = 24 * 60 * 60  # 24시간

ITEM_LIST_PATH = "/OAPI/V2/InventoryBasic/GetBasicProductsList"


def _fetch_items_raw(sess: dict) -> list:
    url = f"{sess['base_url']}{ITEM_LIST_PATH}?SESSION_ID={sess['session_id']}"
    resp = requests.post(url, json={}, timeout=TIMEOUT)
    resp.raise_for_status()
    data = resp.json()

    status = data.get("Status")
    if str(status) != "200":
        error = data.get("Error") or data.get("Errors")
        raise RuntimeError(f"품목조회 실패 Status={status} Error={error}")

    raw_data = data.get("Data", {})
    if isinstance(raw_data, dict):
        result = raw_data.get("Result", [])
        if isinstance(result, str):
            result = json.loads(result)
        return result if isinstance(result, list) else []
    return []


def fetch_item_master(force: bool = False) -> dict:
    """
    전체 품목 마스터를 반환한다.
    반환: {
        prod_cd: {
            "prod_des": str,   품목명
            "brand": str,      CONT1 (브랜드)
            "in_price": float, 입고단가
            "bar_code": str,   바코드
        }
    }
    """
    now = int(time.time())

    # 캐시 확인
    if not force and ITEM_CACHE_FILE.exists():
        try:
            with open(ITEM_CACHE_FILE, "r", encoding="utf-8") as f:
                cache = json.load(f)
            if cache.get("fetched_at", 0) + ITEM_CACHE_TTL > now:
                print(f"  품목 마스터 캐시 사용 ({len(cache['items'])}개, "
                      f"{int((now - cache['fetched_at']) / 60)}분 전 갱신)")
                return cache["items"], cache.get("barcode_index", {})
        except Exception:
            pass

    print("  품목 마스터 조회 중 (전체)...", end=" ", flush=True)
    sess = get_session()
    rows = _fetch_items_raw(sess)
    print(f"{len(rows)}개")

    items = {}
    barcode_index = {}  # BAR_CODE → item info (바코드로 등록된 품목코드 대응용)
    for row in rows:
        if not isinstance(row, dict):
            continue
        prod_cd = str(row.get("PROD_CD") or "").strip()
        if not prod_cd:
            continue

        try:
            in_price = float(str(row.get("IN_PRICE") or 0).replace(",", ""))
        except (ValueError, TypeError):
            in_price = 0.0

        info = {
            "prod_des": str(row.get("PROD_DES") or "").strip(),
            "brand":    str(row.get("CONT1") or "").strip(),
            "in_price": in_price,
            "bar_code": str(row.get("BAR_CODE") or "").strip(),
        }
        items[prod_cd] = info

        # 바코드가 있으면 역방향 인덱스 추가
        bar_code = info["bar_code"]
        if bar_code:
            barcode_index[bar_code] = info

    # 캐시 저장 (바코드 인덱스 포함)
    with open(ITEM_CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump({"fetched_at": now, "items": items, "barcode_index": barcode_index},
                  f, ensure_ascii=False)

    return items, barcode_index
