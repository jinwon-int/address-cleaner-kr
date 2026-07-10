"""clients.py 목킹 테스트: API 응답 파싱, 재시도, 에러코드 처리.

네트워크는 requests.get / session.get 패치로 차단하고,
재시도 테스트는 time.sleep을 패치해 실행 시간을 0으로 유지한다.
픽스처는 실제 응답 형태를 본뜬 익명화 JSON/XML 상수다.
"""

from __future__ import annotations

import pytest
import requests

from address_cleaner import clients
from address_cleaner.clients import (
    JusoClient,
    KoreaPostRoadNameClient,
    juso_key_from_env,
    request_juso,
)

JUSO_OK_RESPONSE = {
    "results": {
        "common": {"errorCode": "0", "errorMessage": "정상", "totalCount": "1"},
        "juso": [
            {
                "roadAddr": "경기도 파주시 하우3길 22",
                "jibunAddr": "경기도 파주시 야당동 57-17",
                "zipNo": "10911",
            }
        ],
    }
}

JUSO_ERROR_RESPONSE = {
    "results": {
        "common": {"errorCode": "E0014", "errorMessage": "승인되지 않은 KEY 입니다."}
    }
}

EPOST_OK_XML = (
    "<NewAddressListResponse><cmmMsgHeader>"
    "<returnCode>00</returnCode><totalCount>1</totalCount>"
    "</cmmMsgHeader><newAddressListAreaCd>"
    "<zipNo>10911</zipNo><lnmAdres>경기도 파주시 하우3길 22</lnmAdres>"
    "</newAddressListAreaCd></NewAddressListResponse>"
)

EPOST_ITEM_FALLBACK_XML = (
    "<response><header><returnCode>00</returnCode></header>"
    "<body><totalCnt>1</totalCnt><items><item>"
    "<zipNo>10911</zipNo><lnmAdres>경기도 파주시 하우3길 22</lnmAdres>"
    "</item></items></body></response>"
)

EPOST_ERROR_XML = (
    "<response><cmmMsgHeader><returnCode>30</returnCode>"
    "<returnMessage>등록되지 않은 서비스키</returnMessage></cmmMsgHeader></response>"
)


class _Response:
    def __init__(self, json_data=None, text="", status_error=False):
        self._json = json_data
        self.text = text
        self._status_error = status_error

    def raise_for_status(self):
        if self._status_error:
            raise requests.HTTPError("500 Server Error")

    def json(self):
        if self._json is None:
            raise ValueError("No JSON object could be decoded")
        return self._json


def _no_sleep(monkeypatch):
    sleeps: list[float] = []
    monkeypatch.setattr(clients.time, "sleep", lambda s: sleeps.append(s))
    return sleeps


# --- request_juso ---


def test_request_juso_parses_total_and_rows(monkeypatch):
    monkeypatch.setattr(
        clients.requests, "get", lambda *a, **k: _Response(JUSO_OK_RESPONSE)
    )

    result = request_juso("key", "경기도 파주시 야당동 57-17")

    assert result["total"] == 1
    assert result["rows"][0]["zipNo"] == "10911"
    assert "error_code" not in result


def test_request_juso_returns_error_fields_without_raising(monkeypatch):
    monkeypatch.setattr(
        clients.requests, "get", lambda *a, **k: _Response(JUSO_ERROR_RESPONSE)
    )

    result = request_juso("bad-key", "경기도 파주시 야당동 57-17")

    assert result["total"] == 0
    assert result["rows"] == []
    assert result["error_code"] == "E0014"
    assert "승인되지 않은" in result["error_message"]


def test_request_juso_retries_transient_errors_with_backoff(monkeypatch):
    sleeps = _no_sleep(monkeypatch)
    calls = {"n": 0}

    def flaky_get(*args, **kwargs):
        calls["n"] += 1
        if calls["n"] <= 2:
            raise requests.ConnectionError("일시적 네트워크 오류")
        return _Response(JUSO_OK_RESPONSE)

    monkeypatch.setattr(clients.requests, "get", flaky_get)

    result = request_juso("key", "경기도 파주시 야당동 57-17")

    assert result["total"] == 1
    assert calls["n"] == 3
    assert sleeps == [1, 2]  # 지수 백오프 2**0, 2**1


def test_request_juso_raises_after_retries_exhausted(monkeypatch):
    _no_sleep(monkeypatch)

    def always_fail(*args, **kwargs):
        raise requests.ConnectionError("network down")

    monkeypatch.setattr(clients.requests, "get", always_fail)

    with pytest.raises(requests.ConnectionError):
        request_juso("key", "경기도 파주시 야당동 57-17", retries=2)


def test_request_juso_retries_non_json_response(monkeypatch):
    # 게이트웨이가 HTML 오류 페이지를 돌려주는 경우: json() ValueError도 재시도 대상.
    _no_sleep(monkeypatch)
    calls = {"n": 0}

    def html_then_json(*args, **kwargs):
        calls["n"] += 1
        if calls["n"] == 1:
            return _Response(json_data=None, text="<html>Bad Gateway</html>")
        return _Response(JUSO_OK_RESPONSE)

    monkeypatch.setattr(clients.requests, "get", html_then_json)

    result = request_juso("key", "경기도 파주시 야당동 57-17")

    assert result["total"] == 1
    assert calls["n"] == 2


def test_request_juso_uses_injected_session(monkeypatch):
    class _Session:
        def __init__(self):
            self.calls = 0

        def get(self, *args, **kwargs):
            self.calls += 1
            return _Response(JUSO_OK_RESPONSE)

    def forbid_module_get(*args, **kwargs):
        raise AssertionError("session이 있으면 requests.get을 쓰면 안 된다")

    monkeypatch.setattr(clients.requests, "get", forbid_module_get)
    session = _Session()

    result = request_juso("key", "경기도 파주시 야당동 57-17", session=session)

    assert result["total"] == 1
    assert session.calls == 1


# --- JusoClient ---


def test_juso_client_requires_key(monkeypatch):
    for name in clients.JUSO_KEY_ENV_VARS:
        monkeypatch.delenv(name, raising=False)

    with pytest.raises(RuntimeError, match="JUSO_CONFIRM_KEY"):
        JusoClient().search("경기도 파주시 야당동 57-17")


def test_juso_client_reads_any_supported_env_var(monkeypatch):
    for name in clients.JUSO_KEY_ENV_VARS:
        monkeypatch.delenv(name, raising=False)
    monkeypatch.setenv("JUSO_API_KEY", "env-key")

    assert juso_key_from_env() == "env-key"
    assert JusoClient().key == "env-key"


def test_juso_client_maps_api_error_to_search_result(monkeypatch):
    monkeypatch.setattr(
        clients.requests, "get", lambda *a, **k: _Response(JUSO_ERROR_RESPONSE)
    )

    result = JusoClient(key="bad-key").search("경기도 파주시 야당동 57-17")

    assert result.provider == "juso"
    assert result.has_error
    assert not result.found
    assert result.first["errorCode"] == "E0014"


def test_juso_client_returns_first_row(monkeypatch):
    monkeypatch.setattr(
        clients.requests, "get", lambda *a, **k: _Response(JUSO_OK_RESPONSE)
    )

    result = JusoClient(key="key").search("경기도 파주시 야당동 57-17")

    assert result.found
    assert not result.has_error
    assert result.total_count == 1
    assert result.first["roadAddr"] == "경기도 파주시 하우3길 22"


# --- KoreaPostRoadNameClient ---


def test_epost_requires_key(monkeypatch):
    for name in clients.EPOST_KEY_ENV_VARS:
        monkeypatch.delenv(name, raising=False)

    with pytest.raises(RuntimeError, match="EPOST_SERVICE_KEY"):
        KoreaPostRoadNameClient().search("하우3길 22")


def test_epost_parses_total_count_and_first_item(monkeypatch):
    monkeypatch.setattr(
        clients.requests, "get", lambda *a, **k: _Response(text=EPOST_OK_XML)
    )

    result = KoreaPostRoadNameClient(key="key").search("하우3길 22")

    assert result.total_count == 1
    assert result.first["zipNo"] == "10911"
    assert not result.has_error


def test_epost_falls_back_to_item_element(monkeypatch):
    monkeypatch.setattr(
        clients.requests,
        "get",
        lambda *a, **k: _Response(text=EPOST_ITEM_FALLBACK_XML),
    )

    result = KoreaPostRoadNameClient(key="key").search("하우3길 22")

    assert result.total_count == 1  # totalCnt 폴백
    assert result.first["lnmAdres"] == "경기도 파주시 하우3길 22"


def test_epost_marks_error_return_code(monkeypatch):
    monkeypatch.setattr(
        clients.requests, "get", lambda *a, **k: _Response(text=EPOST_ERROR_XML)
    )

    result = KoreaPostRoadNameClient(key="bad-key").search("하우3길 22")

    assert result.has_error
    assert result.first["returnCode"] == "30"
    assert "서비스키" in result.first["returnMessage"]


def test_epost_retries_broken_xml(monkeypatch):
    sleeps = _no_sleep(monkeypatch)
    calls = {"n": 0}

    def broken_then_ok(*args, **kwargs):
        calls["n"] += 1
        if calls["n"] == 1:
            return _Response(text="<broken><xml")
        return _Response(text=EPOST_OK_XML)

    monkeypatch.setattr(clients.requests, "get", broken_then_ok)

    result = KoreaPostRoadNameClient(key="key").search("하우3길 22")

    assert result.total_count == 1
    assert calls["n"] == 2
    assert sleeps == [1]


def test_epost_raises_after_retries_exhausted(monkeypatch):
    _no_sleep(monkeypatch)
    monkeypatch.setattr(
        clients.requests, "get", lambda *a, **k: _Response(text="<broken><xml")
    )

    with pytest.raises(Exception):
        KoreaPostRoadNameClient(key="key").search("하우3길 22", retries=1)
