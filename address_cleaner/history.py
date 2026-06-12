"""검증 이력 영속화(SQLite).

월 배치마다 같은 물건 주소가 반복되므로, 검증 결과를 실행 단위가 아니라
파일 DB에 쌓아 ① 최근 결과 재사용으로 API 호출을 줄이고
② "지난달 1건 → 이번달 0건" 같은 판정 변화(행정구역 개편·건물 멸실 신호)를 감지한다.
"""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path


@dataclass(frozen=True)
class HistoryEntry:
    verdict: str
    detail: str
    checked_at: str


class VerifyHistory:
    def __init__(self, path: str | Path, max_age_days: int = 14):
        self.max_age_days = max_age_days
        self.conn = sqlite3.connect(str(path))
        self.conn.execute(
            "CREATE TABLE IF NOT EXISTS verify_history ("
            "query TEXT NOT NULL, kind TEXT NOT NULL, verdict TEXT NOT NULL, "
            "detail TEXT NOT NULL DEFAULT '', checked_at TEXT NOT NULL)"
        )
        self.conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_verify_history ON verify_history(query, kind, checked_at)"
        )
        self.conn.commit()

    def latest(self, query: str, kind: str) -> HistoryEntry | None:
        row = self.conn.execute(
            "SELECT verdict, detail, checked_at FROM verify_history "
            "WHERE query = ? AND kind = ? ORDER BY checked_at DESC LIMIT 1",
            (query, kind),
        ).fetchone()
        return HistoryEntry(*row) if row else None

    def fresh(self, query: str, kind: str) -> HistoryEntry | None:
        """max_age_days 내에 검증한 적 있으면 그 결과 (API 호출 생략용)."""
        entry = self.latest(query, kind)
        if entry is None:
            return None
        try:
            checked = datetime.fromisoformat(entry.checked_at)
        except ValueError:
            return None
        if checked < datetime.now() - timedelta(days=self.max_age_days):
            return None
        return entry

    def record(self, query: str, kind: str, verdict: str, detail: str, checked_at: str | None = None) -> None:
        self.conn.execute(
            "INSERT INTO verify_history VALUES (?, ?, ?, ?, ?)",
            (query, kind, verdict, detail, checked_at or datetime.now().isoformat(timespec="seconds")),
        )
        self.conn.commit()

    def close(self) -> None:
        self.conn.close()
