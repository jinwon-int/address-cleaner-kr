"""history.py 테스트: 이력 왕복, 신선도 판정, 파싱 불가 행 처리.

시계를 조작하지 않고 checked_at을 과거로 직접 기록해 만료를 테스트한다.
"""

from __future__ import annotations

from datetime import datetime, timedelta

from address_cleaner.history import VerifyHistory

QUERY = "경기도 파주시 야당동 57-17 정우펠리스 제303동 제101호"


def test_record_then_latest_roundtrip(tmp_path):
    history = VerifyHistory(tmp_path / "history.sqlite")
    history.record(QUERY, "lot", "verified", "JUSO[상세포함] 1건")

    entry = history.latest(QUERY, "lot")
    history.close()

    assert entry is not None
    assert entry.verdict == "verified"
    assert entry.detail == "JUSO[상세포함] 1건"
    assert entry.checked_at


def test_latest_is_scoped_by_query_and_kind(tmp_path):
    history = VerifyHistory(tmp_path / "history.sqlite")
    history.record(QUERY, "lot", "verified", "")

    assert history.latest(QUERY, "road") is None
    assert history.latest("다른 주소", "lot") is None
    history.close()


def test_fresh_reuses_entry_within_max_age(tmp_path):
    history = VerifyHistory(tmp_path / "history.sqlite", max_age_days=14)
    checked_at = (datetime.now() - timedelta(days=13)).isoformat(timespec="seconds")
    history.record(QUERY, "lot", "verified", "", checked_at=checked_at)

    entry = history.fresh(QUERY, "lot")
    history.close()

    assert entry is not None
    assert entry.verdict == "verified"


def test_fresh_expires_entry_older_than_max_age(tmp_path):
    history = VerifyHistory(tmp_path / "history.sqlite", max_age_days=14)
    checked_at = (datetime.now() - timedelta(days=15)).isoformat(timespec="seconds")
    history.record(QUERY, "lot", "verified", "", checked_at=checked_at)

    assert history.fresh(QUERY, "lot") is None
    # 만료돼도 판정 변화 감지용 latest는 남는다.
    assert history.latest(QUERY, "lot") is not None
    history.close()


def test_fresh_treats_unparseable_checked_at_as_stale(tmp_path):
    history = VerifyHistory(tmp_path / "history.sqlite")
    history.record(QUERY, "lot", "verified", "", checked_at="지난달 어느 날")

    assert history.fresh(QUERY, "lot") is None
    history.close()


def test_latest_returns_newest_record_for_verdict_change_detection(tmp_path):
    # excel.py의 판정 변경 감지 시나리오: 지난달 verified → 이번달 missing.
    history = VerifyHistory(tmp_path / "history.sqlite")
    history.record(QUERY, "lot", "verified", "", checked_at="2026-05-10T09:00:00")
    history.record(QUERY, "lot", "missing", "", checked_at="2026-06-10T09:00:00")

    entry = history.latest(QUERY, "lot")
    history.close()

    assert entry is not None
    assert entry.verdict == "missing"
    assert entry.checked_at.startswith("2026-06-10")
