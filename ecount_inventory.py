"""
이카운트(ECOUNT) 재고현황 조회 모듈
─────────────────────────────────────────
창고별 재고현황 API (GetListInventoryBalanceStatusByLocation) 를 호출한다.
온라인사업부 가용재고는 창고코드 20002(디아이 판매) 를 본다.
"""

import json
import requests

from ecount_auth import get_session

TIMEOUT = 30

INVENTORY_BY_LOCATION_PATH = "/OAPI/V2/InventoryBalance/GetListInventoryBalanceStatusByLocation"


def get_inventory_by_location(base_date: str, wh_cd: str = None,
                              prod_cd: str = None, force_login: bool = False) -> dict:
    """
    창고별 재고현황을 조회해 '원본 응답(JSON)' 을 그대로 dict 로 반환한다.

    base_date : 기준일자 'YYYYMMDD'
    wh_cd     : 창고코드 (예: '20002'). None 이면 전체 창고.
    prod_cd   : 품목코드. None 이면 전체 품목.
    """
    sess = get_session(force=force_login)
    url = f"{sess['base_url']}{INVENTORY_BY_LOCATION_PATH}?SESSION_ID={sess['session_id']}"

    body = {"BASE_DATE": base_date}
    if wh_cd:
        body["WH_CD"] = wh_cd
    if prod_cd:
        body["PROD_CD"] = prod_cd

    resp = requests.post(url, json=body, timeout=TIMEOUT)
    resp.raise_for_status()
    return resp.json()


def extract_rows(raw: dict) -> list:
    """이카운트 응답 래퍼에서 실제 재고 행 리스트를 방어적으로 꺼낸다."""
    data = raw.get("Data")
    if isinstance(data, dict):
        for key in ("Result", "Datas", "Data", "ResultData"):
            val = data.get(key)
            if isinstance(val, list):
                return val
    if isinstance(data, list):
        return data
    return []
