"""
이카운트(ECOUNT) 설정 로더
.env 파일에서 ECOUNT_* 값을 읽어 반환한다.

.env 에 필요한 항목:
  ECOUNT_COM_CODE      = 회사코드 (이카운트 로그인 시 쓰는 회사코드)
  ECOUNT_USER_ID       = 사용자 ID (API용 사용자)
  ECOUNT_API_CERT_KEY  = 발급받은 API 인증키
  ECOUNT_IS_TEST       = true  → 테스트(sboapi) 서버 / false → 실(oapi) 서버
"""

import os
from dataclasses import dataclass
from dotenv import load_dotenv

load_dotenv()


@dataclass
class EcountConfig:
    com_code: str
    user_id: str
    api_cert_key: str
    is_test: bool

    def __repr__(self):
        key_preview = (self.api_cert_key[:6] + "...") if self.api_cert_key else "(없음)"
        server = "테스트(sboapi)" if self.is_test else "실서버(oapi)"
        return (f"EcountConfig(com_code='{self.com_code}', user_id='{self.user_id}', "
                f"api_cert_key='{key_preview}', server={server})")


def load_ecount_config() -> EcountConfig:
    com_code = os.getenv("ECOUNT_COM_CODE", "").strip()
    user_id = os.getenv("ECOUNT_USER_ID", "").strip()
    api_cert_key = os.getenv("ECOUNT_API_CERT_KEY", "").strip()
    is_test = os.getenv("ECOUNT_IS_TEST", "false").strip().lower() in ("1", "true", "yes", "y")

    missing = []
    if not com_code:
        missing.append("ECOUNT_COM_CODE")
    if not user_id:
        missing.append("ECOUNT_USER_ID")
    if not api_cert_key:
        missing.append("ECOUNT_API_CERT_KEY")
    if missing:
        raise RuntimeError(
            ".env 에 다음 항목이 비어 있습니다: " + ", ".join(missing)
        )

    return EcountConfig(
        com_code=com_code,
        user_id=user_id,
        api_cert_key=api_cert_key,
        is_test=is_test,
    )
