"""공개 라이브러리 API(#15) 테스트: 최상위 export, verify_address, py.typed."""

from __future__ import annotations

from pathlib import Path

import pytest

import address_cleaner
from address_cleaner import (
    JusoClient,
    KoreaPostRoadNameClient,
    NormalizedAddress,
    SearchResult,
    VerifyResult,
    normalize_for_search,
    process_workbook,
    verify_address,
)
from address_cleaner.clients import EPOST_KEY_ENV_VARS, JUSO_KEY_ENV_VARS


def test_top_level_exports_are_usable():
    assert callable(normalize_for_search)
    assert callable(verify_address)
    assert callable(process_workbook)
    assert {JusoClient, KoreaPostRoadNameClient, NormalizedAddress, SearchResult}
    assert set(address_cleaner.__all__) >= {
        "normalize_for_search",
        "verify_address",
        "VerifyResult",
        "__version__",
    }


def test_version_comes_from_package_metadata():
    # 버전 단일 출처는 pyproject.toml — import 가능한 문자열이어야 한다.
    assert isinstance(address_cleaner.__version__, str)
    assert address_cleaner.__version__


def test_py_typed_marker_ships_with_package():
    marker = Path(address_cleaner.__file__).parent / "py.typed"
    assert marker.exists()


def test_verify_address_returns_dataclass():
    class _FakeJuso:
        key = "test-key"

        def search(self, query: str, count: int = 5):
            return SearchResult(
                "juso",
                1,
                {"roadAddr": "경기도 파주시 하우3길 22", "zipNo": "10911"},
                {},
            )

    result = verify_address(
        "경기도 파주시 야당동 57-17", "lot", juso=_FakeJuso(), epost=None
    )

    assert isinstance(result, VerifyResult)
    assert result.verdict == "verified"
    assert result.verified
    assert "10911" in result.detail
    assert result.correction is None


def test_verify_address_without_any_key_raises(monkeypatch):
    for name in JUSO_KEY_ENV_VARS + EPOST_KEY_ENV_VARS:
        monkeypatch.delenv(name, raising=False)

    with pytest.raises(RuntimeError, match="API key"):
        verify_address("경기도 파주시 야당동 57-17", "lot")


def test_verify_address_harvests_correction():
    class _BaseOnlyJuso:
        key = "test-key"

        def search(self, query: str, count: int = 5):
            if query == "경기도 파주시 야당동 57-17":
                return SearchResult("juso", 1, {"roadAddr": "표준", "zipNo": "1"}, {})
            return SearchResult("juso", 0, {}, {})

    result = verify_address(
        "경기도 파주시 야당동 57-17 정우펠리스 제101호",
        "lot",
        juso=_BaseOnlyJuso(),
    )

    assert result.verified
    assert result.correction is not None
    assert result.correction["type"] == "상세부제거"
