"""한국 주소 정제 + juso.go.kr/우체국 API 검증 라이브러리.

>>> from address_cleaner import normalize_for_search, verify_address
>>> normalized = normalize_for_search("경기도 파주시 야당동 57-17 정우펠리스 제1층 제101호")
>>> normalized.query
'경기도 파주시 야당동 57-17 정우펠리스 제101호'
>>> verify_address(normalized.query, normalized.kind).verdict  # API 키 필요
'verified'
"""

from importlib.metadata import PackageNotFoundError, version

from .clients import JusoClient, KoreaPostRoadNameClient, SearchResult
from .excel import VerifyResult, process_workbook, verify_address
from .normalizer import NormalizedAddress, normalize_for_search

try:
    # 버전 단일 출처는 pyproject.toml — 설치 메타데이터에서 읽는다.
    __version__ = version("address-cleaner-kr")
except PackageNotFoundError:  # 미설치 소스 트리에서 import한 경우
    __version__ = "0+unknown"

__all__ = [
    "JusoClient",
    "KoreaPostRoadNameClient",
    "NormalizedAddress",
    "SearchResult",
    "VerifyResult",
    "__version__",
    "normalize_for_search",
    "process_workbook",
    "verify_address",
]
