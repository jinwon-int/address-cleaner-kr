from __future__ import annotations

from dataclasses import dataclass
import os
import time
import xml.etree.ElementTree as ET
from typing import Any

import requests

JUSO_ENDPOINT = "https://business.juso.go.kr/addrlink/addrLinkApi.do"

# 일반 정제 모드와 등기소 모드가 서로 다른 이름을 써 왔어서 둘 다 허용한다.
JUSO_KEY_ENV_VARS = ("JUSO_CONFIRM_KEY", "JUSO_CONFM_KEY", "JUSO_API_KEY", "CONFM_KEY")
EPOST_KEY_ENV_VARS = ("EPOST_SERVICE_KEY", "KOREAPOST_SERVICE_KEY")


def juso_key_from_env() -> str | None:
    for name in JUSO_KEY_ENV_VARS:
        value = os.getenv(name)
        if value:
            return value
    return None


def request_juso(
    key: str,
    keyword: str,
    count: int = 10,
    *,
    timeout: float = 15.0,
    session: requests.Session | None = None,
    retries: int = 3,
) -> dict[str, Any]:
    """Juso API 한 번 호출. 일시적 네트워크 오류/비정상 응답은 지수 백오프로 재시도한다."""
    params = {
        "confmKey": key,
        "currentPage": "1",
        "countPerPage": str(count),
        "keyword": keyword,
        "resultType": "json",
    }
    get = session.get if session is not None else requests.get
    data: dict[str, Any] = {}
    for attempt in range(retries + 1):
        try:
            response = get(JUSO_ENDPOINT, params=params, timeout=timeout)
            response.raise_for_status()
            data = response.json()
            break
        except (requests.RequestException, ValueError):
            if attempt == retries:
                raise
            time.sleep(2**attempt)
    common = data.get("results", {}).get("common", {})
    code = str(common.get("errorCode", ""))
    if code not in {"0", "00"}:
        return {
            "total": 0,
            "rows": [],
            "error_code": code,
            "error_message": common.get("errorMessage", ""),
            "raw": data,
        }
    return {
        "total": int(common.get("totalCount") or 0),
        "rows": data.get("results", {}).get("juso") or [],
        "raw": data,
    }


@dataclass
class SearchResult:
    provider: str
    total_count: int
    first: dict
    raw: dict | str

    @property
    def found(self) -> bool:
        return self.total_count > 0

    @property
    def has_error(self) -> bool:
        return "errorCode" in self.first or "returnCode" in self.first


class JusoClient:
    endpoint = JUSO_ENDPOINT

    def __init__(
        self,
        key: str | None = None,
        timeout: float = 5.0,
        session: requests.Session | None = None,
    ):
        self.key = key or juso_key_from_env()
        self.timeout = timeout
        # 병렬 검증에서 커넥션을 재사용하도록 Session 주입을 지원한다 (기본은 단발 요청).
        self.session = session

    def search(self, keyword: str, count: int = 10) -> SearchResult:
        if not self.key:
            raise RuntimeError("JUSO_CONFIRM_KEY is required for juso.go.kr validation")
        result = request_juso(
            self.key, keyword, count, timeout=self.timeout, session=self.session
        )
        if "error_code" in result:
            return SearchResult(
                "juso",
                0,
                {
                    "errorCode": result["error_code"],
                    "errorMessage": result["error_message"],
                },
                result["raw"],
            )
        rows = result["rows"]
        return SearchResult(
            "juso", result["total"], rows[0] if rows else {}, result["raw"]
        )


class KoreaPostRoadNameClient:
    endpoint = (
        "http://openapi.epost.go.kr:80/postal/retrieveNewAdressAreaCdService/"
        "retrieveNewAdressAreaCdService/getNewAddressListAreaCd"
    )

    def __init__(
        self,
        key: str | None = None,
        timeout: float = 5.0,
        session: requests.Session | None = None,
    ):
        self.key = (
            next(
                (os.getenv(name) for name in EPOST_KEY_ENV_VARS if os.getenv(name)),
                None,
            )
            if key is None
            else key
        )
        self.timeout = timeout
        self.session = session

    def search(
        self, keyword: str, search_se: str = "road", count: int = 10, retries: int = 3
    ) -> SearchResult:
        if not self.key:
            raise RuntimeError(
                "EPOST_SERVICE_KEY is required for Korea Post validation"
            )
        params = {
            "ServiceKey": self.key,
            "searchSe": search_se,
            "srchwrd": keyword,
            "countPerPage": str(count),
            "currentPage": "1",
        }
        get = self.session.get if self.session is not None else requests.get
        # Juso 쪽 request_juso와 같은 기준으로 일시적 네트워크 오류/깨진 응답을 재시도한다.
        for attempt in range(retries + 1):
            try:
                response = get(self.endpoint, params=params, timeout=self.timeout)
                response.raise_for_status()
                text = response.text
                root = ET.fromstring(text)
                break
            except (requests.RequestException, ET.ParseError):
                if attempt == retries:
                    raise
                time.sleep(2**attempt)
        return_code = _text(root, ".//returnCode")
        if return_code and return_code != "00":
            return SearchResult(
                "epost",
                0,
                {
                    "returnCode": return_code,
                    "returnMessage": _text(root, ".//returnMessage"),
                },
                text,
            )
        total = int(_text(root, ".//totalCount") or _text(root, ".//totalCnt") or "0")
        first = {}
        item = root.find(".//newAddressListAreaCd")
        if item is None:
            item = root.find(".//item")
        if item is not None:
            first = {child.tag: child.text for child in list(item)}
        return SearchResult("epost", total, first, text)


def _text(root: ET.Element, path: str) -> str:
    value = root.findtext(path)
    return value.strip() if value else ""
