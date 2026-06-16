"""
이카운트(ECOUNT) 연결 테스트
─────────────────────────────────────────
실행하면 아래 순서로 접속을 확인하고, 어디서 막혔는지 한국어로 보여줍니다.
  1) .env 설정 읽기
  2) Zone 조회
  3) 로그인 → SESSION_ID 발급
세 가지가 모두 통과하면 이카운트 API 연결 준비가 끝난 것입니다.
"""

import sys
import traceback

# 윈도우 콘솔에서 한글/이모지 출력이 깨지거나 멈추지 않도록 UTF-8 로 강제
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from ecount_config import load_ecount_config
from ecount_auth import lookup_zone, login


def main():
    print("=" * 50)
    print(" 이카운트(ECOUNT) API 연결 테스트")
    print("=" * 50)

    try:
        cfg = load_ecount_config()
    except Exception as e:
        print(f"\n[1단계 실패] 설정을 읽지 못했습니다.\n  → {e}")
        print("\n.env 파일에 ECOUNT_COM_CODE / ECOUNT_USER_ID / ECOUNT_API_CERT_KEY 를 채워주세요.")
        return 1

    print(f"\n[1단계 OK] 설정 읽기 완료")
    print(f"  {cfg!r}")

    try:
        zone = lookup_zone(cfg.com_code, cfg.is_test)
    except Exception as e:
        print(f"\n[2단계 실패] Zone 조회 실패")
        print(f"  → {e}")
        print("  ※ 회사코드(ECOUNT_COM_CODE)가 맞는지, 테스트/실서버 구분(ECOUNT_IS_TEST)이 맞는지 확인하세요.")
        return 1

    print(f"\n[2단계 OK] Zone 조회 완료 → ZONE = {zone}")

    try:
        session_id = login(cfg.com_code, cfg.user_id, cfg.api_cert_key, zone, cfg.is_test)
    except Exception as e:
        print(f"\n[3단계 실패] 로그인 실패")
        print(f"  → {e}")
        print("  ※ 사용자ID 와 인증키를 확인하세요.")
        print("  ※ 실서버(oapi)는 '등록된 IP'에서만 접속됩니다. 이 PC의 IP가 이카운트 [IP등록]에 있는지 확인하세요.")
        print("  ※ 실서버는 로그인이 10분에 1회 제한이라, 방금 실패했다면 잠시 후 다시 시도하세요.")
        return 1

    masked = session_id[:6] + "..." + session_id[-4:] if len(session_id) > 12 else "(발급됨)"
    print(f"\n[3단계 OK] 로그인 성공 → SESSION_ID = {masked}")

    print("\n" + "=" * 50)
    print(" ✅ 연결 성공! 이카운트 API 사용 준비가 끝났습니다.")
    print("=" * 50)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        print("\n[예상치 못한 오류]")
        traceback.print_exc()
        sys.exit(1)
