from __future__ import annotations

from dataclasses import dataclass
import os
import xml.etree.ElementTree as ET

import requests


@dataclass
class SearchResult:
    provider: str
    total_count: int
    first: dict
    raw: dict | str

    @property
    def found(self) -> bool:
        return self.total_count > 0


class JusoClient:
    endpoint = "https://business.juso.go.kr/addrlink/addrLinkApi.do"

    def __init__(self, key: str | None = None, timeout: float = 5.0):
        self.key = key or os.getenv("JUSO_CONFIRM_KEY") or os.getenv("JUSO_API_KEY")
        self.timeout = timeout

    def search(self, keyword: str, count: int = 10) -> SearchResult:
        if not self.key:
            raise RuntimeError("JUSO_CONFIRM_KEY is required for juso.go.kr validation")
        params = {
            "confmKey": self.key,
            "currentPage": 1,
            "countPerPage": count,
            "keyword": keyword,
            "resultType": "json",
        }
        response = requests.get(self.endpoint, params=params, timeout=self.timeout)
        response.raise_for_status()
        data = response.json()
        common = data.get("results", {}).get("common", {})
        error_code = common.get("errorCode")
        if error_code != "0":
            return SearchResult("juso", 0, {"errorCode": error_code, "errorMessage": common.get("errorMessage")}, data)
        total = int(common.get("totalCount") or 0)
        rows = data.get("results", {}).get("juso") or []
        return SearchResult("juso", total, rows[0] if rows else {}, data)


class KoreaPostRoadNameClient:
    endpoint = (
        "http://openapi.epost.go.kr:80/postal/retrieveNewAdressAreaCdService/"
        "retrieveNewAdressAreaCdService/getNewAddressListAreaCd"
    )

    def __init__(self, key: str | None = None, timeout: float = 5.0):
        self.key = key or os.getenv("EPOST_SERVICE_KEY") or os.getenv("KOREAPOST_SERVICE_KEY")
        self.timeout = timeout

    def search(self, keyword: str, search_se: str = "road", count: int = 10) -> SearchResult:
        if not self.key:
            raise RuntimeError("EPOST_SERVICE_KEY is required for Korea Post validation")
        params = {
            "ServiceKey": self.key,
            "searchse": search_se,
            "srchwrd": keyword,
            "countperpage": count,
            "currentpage": 1,
        }
        response = requests.get(self.endpoint, params=params, timeout=self.timeout)
        response.raise_for_status()
        text = response.text
        root = ET.fromstring(text)
        return_code = _text(root, ".//returnCode")
        if return_code and return_code != "00":
            return SearchResult(
                "epost",
                0,
                {"returnCode": return_code, "returnMessage": _text(root, ".//returnMessage")},
                text,
            )
        total = int(_text(root, ".//totalCount") or _text(root, ".//totalCnt") or "0")
        first = {}
        item = root.find(".//newAddressListAreaCd") or root.find(".//item")
        if item is not None:
            first = {child.tag: child.text for child in list(item)}
        return SearchResult("epost", total, first, text)


def _text(root: ET.Element, path: str) -> str:
    value = root.findtext(path)
    return value.strip() if value else ""
