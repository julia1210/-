"""
이카운트(ECOUNT) OpenAPI 인증 모듈
─────────────────────────────────────────
이카운트 OpenAPI는 3단계로 동작합니다.
  1) Zone 조회   : 회사코드(COM_CODE)로 우리 회사가 속한 서버(Zone)를 확인
  2) 로그인      : 회사코드 + 사용자ID + 인증키 + Zone  →  SESSION_ID 발급
  3) 실제 API    : SESSION_ID 를 URL 파라미터로 붙여 재고현황 등 조회

SESSION_ID 는 일정 시간이 지나면 만료되므로 .ecount_session.json 에 캐싱하고,
만료 임박/실패 시 자동으로 다시 로그인합니다.

※ 실서버(oapi)는 로그인이 '10분에 1회'로 제한되므로 세션을 충분히 재사용합니다.
"""

import json
import time
import requests
from pathlib import Path

from ecount_config import load_ecount_config

# 세션 캐시 파일 (자동 생성)
SESSION_CACHE_FILE = Path(".ecount_session.json")

# SESSION_ID 를 안전하게 재사용할 시간(초). 실서버 로그인 한도(10분/회)를 고려해 넉넉히 25분.
SESSION_TTL_SECONDS = 25 * 60

TIMEOUT = 30


def _base_domain(zone: str, is_test: bool) -> str:
    prefix = "sboapi" if is_test else "oapi"
    return f"https://{prefix}{zone}.ecount.com"


def _zone_domain(is_test: bool) -> str:
    prefix = "sboapi" if is_test else "oapi"
    return f"https://{prefix}.ecount.com"


def lookup_zone(com_code: str, is_test: bool) -> str:
    """1단계: 회사코드로 Zone 을 조회한다."""
    url = f"{_zone_domain(is_test)}/OAPI/V2/Zone"
    resp = requests.post(url, json={"COM_CODE": com_code}, timeout=TIMEOUT)
    resp.raise_for_status()
    data = resp.json()
    zone = (data.get("Data") or {}).get("ZONE")
    if not zone:
        raise RuntimeError(f"Zone 조회 실패 — 응답: {json.dumps(data, ensure_ascii=False)}")
    return zone


def login(com_code: str, user_id: str, api_cert_key: str, zone: str,
          is_test: bool, lan_type: str = "ko-KR") -> str:
    """2단계: 로그인하여 SESSION_ID 를 발급받는다."""
    url = f"{_base_domain(zone, is_test)}/OAPI/V2/OAPILogin"
    payload = {
        "COM_CODE": com_code,
        "USER_ID": user_id,
        "API_CERT_KEY": api_cert_key,
        "LAN_TYPE": lan_type,
        "ZONE": zone,
    }
    resp = requests.post(url, json=payload, timeout=TIMEOUT)
    resp.raise_for_status()
    data = resp.json()
    datas = (data.get("Data") or {}).get("Datas") or {}
    session_id = datas.get("SESSION_ID")
    if not session_id:
        raise RuntimeError(f"로그인 실패 — 응답: {json.dumps(data, ensure_ascii=False)}")
    return session_id


def _load_session_cache() -> dict:
    if not SESSION_CACHE_FILE.exists():
        return {}
    try:
        with open(SESSION_CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return {}


def _save_session_cache(cache: dict) -> None:
    with open(SESSION_CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def get_session(force: bool = False) -> dict:
    """
    유효한 세션 정보를 반환한다. (캐시 우선, 없거나 만료 임박이면 새로 로그인)
    반환: {"zone", "session_id", "base_url", "is_test"}
    """
    cfg = load_ecount_config()
    cache = _load_session_cache()
    now = int(time.time())

    if (not force
            and cache.get("com_code") == cfg.com_code
            and cache.get("is_test") == cfg.is_test
            and cache.get("issued_at", 0) + SESSION_TTL_SECONDS > now
            and cache.get("session_id")):
        return {
            "zone": cache["zone"],
            "session_id": cache["session_id"],
            "base_url": _base_domain(cache["zone"], cfg.is_test),
            "is_test": cfg.is_test,
        }

    zone = lookup_zone(cfg.com_code, cfg.is_test)
    session_id = login(cfg.com_code, cfg.user_id, cfg.api_cert_key, zone, cfg.is_test)

    _save_session_cache({
        "com_code": cfg.com_code,
        "is_test": cfg.is_test,
        "zone": zone,
        "session_id": session_id,
        "issued_at": now,
    })
    return {
        "zone": zone,
        "session_id": session_id,
        "base_url": _base_domain(zone, cfg.is_test),
        "is_test": cfg.is_test,
    }
