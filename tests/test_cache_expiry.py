"""registry JSON 캐시 만료(#14) 테스트: v2 포맷, 만료 재호출, 저장 시 청소."""

from __future__ import annotations

import json
from datetime import datetime, timedelta

import pytest

from address_cleaner import juso_search
from address_cleaner.juso_search import (
    DEFAULT_CACHE_MAX_AGE_DAYS,
    cache_entry_fresh,
    juso_query,
    save_cache,
    set_cache_max_age_days,
    stamp_cache_entry,
)

KEYWORD = "경기도 파주시 야당동 57-17"
CACHE_KEY = f"clean:5:{KEYWORD}"


@pytest.fixture(autouse=True)
def _reset_max_age(monkeypatch):
    # 직렬 경로의 0.04초 대기를 없애 테스트 시간을 0으로 유지한다.
    monkeypatch.setattr(juso_search.time, "sleep", lambda s: None)
    yield
    set_cache_max_age_days(DEFAULT_CACHE_MAX_AGE_DAYS)


def _stamped(days_ago: int) -> dict:
    return {
        "keyword": KEYWORD,
        "total": 1,
        "rows": [{"roadAddr": "표준주소"}],
        "cached_at": (datetime.now() - timedelta(days=days_ago)).isoformat(
            timespec="seconds"
        ),
    }


def _count_requests(monkeypatch):
    calls = {"n": 0}

    def fake_request(key, keyword, count, timeout=15, session=None):
        calls["n"] += 1
        return {"total": 1, "rows": [{"roadAddr": "재호출결과"}], "raw": {}}

    monkeypatch.setattr(juso_search, "request_juso", fake_request)
    return calls


def test_fresh_entry_skips_api_call(monkeypatch):
    calls = _count_requests(monkeypatch)
    cache = {CACHE_KEY: _stamped(days_ago=1)}

    res = juso_query(None, "key", KEYWORD, cache)

    assert calls["n"] == 0
    assert res["rows"][0]["roadAddr"] == "표준주소"


def test_expired_entry_is_requeried_and_overwritten(monkeypatch):
    calls = _count_requests(monkeypatch)
    cache = {CACHE_KEY: _stamped(days_ago=15)}

    res = juso_query(None, "key", KEYWORD, cache)

    assert calls["n"] == 1
    assert res["rows"][0]["roadAddr"] == "재호출결과"
    assert cache[CACHE_KEY]["rows"][0]["roadAddr"] == "재호출결과"
    assert "cached_at" in cache[CACHE_KEY]  # v2 포맷으로 덮어쓴다


def test_legacy_entry_without_cached_at_is_treated_as_expired(monkeypatch):
    calls = _count_requests(monkeypatch)
    cache = {CACHE_KEY: {"keyword": KEYWORD, "total": 1, "rows": []}}

    juso_query(None, "key", KEYWORD, cache)

    assert calls["n"] == 1


def test_max_age_zero_keeps_legacy_behavior(monkeypatch):
    calls = _count_requests(monkeypatch)
    set_cache_max_age_days(0)
    cache = {CACHE_KEY: {"keyword": KEYWORD, "total": 1, "rows": []}}

    res = juso_query(None, "key", KEYWORD, cache)

    assert calls["n"] == 0  # 만료 없음: 구버전 엔트리도 그대로 쓴다
    assert res["total"] == 1


def test_save_cache_drops_expired_entries(tmp_path):
    cache_file = tmp_path / "cache.json"
    cache = {
        "fresh": _stamped(days_ago=1),
        "stale": _stamped(days_ago=30),
        "legacy": {"keyword": "구버전", "total": 0, "rows": []},
    }

    save_cache(cache_file, cache)

    saved = json.loads(cache_file.read_text(encoding="utf-8"))
    assert set(saved) == {"fresh"}
    # 메모리 캐시는 그대로 (실행 중 재사용), 파일만 청소된다.
    assert set(cache) == {"fresh", "stale", "legacy"}


def test_save_cache_keeps_everything_when_expiry_disabled(tmp_path):
    set_cache_max_age_days(0)
    cache_file = tmp_path / "cache.json"
    cache = {"legacy": {"keyword": "구버전", "total": 0, "rows": []}}

    save_cache(cache_file, cache)

    saved = json.loads(cache_file.read_text(encoding="utf-8"))
    assert set(saved) == {"legacy"}


def test_stamp_and_fresh_roundtrip():
    entry = stamp_cache_entry({"keyword": KEYWORD, "total": 1, "rows": []})
    assert cache_entry_fresh(entry)
    assert not cache_entry_fresh({"cached_at": "이상한 값"})
    assert not cache_entry_fresh("문자열 엔트리")


def test_registry_cli_accepts_cache_max_age_days(monkeypatch, capsys):
    from address_cleaner.registry.cli import main

    for name in ("JUSO_CONFIRM_KEY", "JUSO_CONFM_KEY", "JUSO_API_KEY", "CONFM_KEY"):
        monkeypatch.delenv(name, raising=False)

    # 인자 파싱은 키 검사보다 먼저이므로, exit 2(키 없음)면 인자가 수용된 것이다.
    exit_code = main(["input.xlsx", "--cache-max-age-days", "7"])

    captured = capsys.readouterr()
    assert exit_code == 2
    assert "JUSO_CONFM_KEY" in captured.err
