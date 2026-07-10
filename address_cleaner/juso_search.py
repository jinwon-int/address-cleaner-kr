"""Juso 검색 공용 인프라 — 등기(registry)·일반(excel) 모드가 공유한다.

레이트리미터, JSON 응답 캐시, 캐시 잠금, 캐시를 경유하는 juso_query를 담는다.
registry/juso.py에서 옮겨 왔고, 기존 경로는 그대로 re-export 된다.
"""

from __future__ import annotations

import json
import threading
import time
from pathlib import Path
from typing import Any

import requests

from .clients import request_juso


class RateLimiter:
    """스레드 간 공유하는 전역 토큰버킷. 여러 워커가 동시에 호출해도
    초당 호출 수를 max_per_sec 이하로 묶어 API 차단/RemoteDisconnected를 줄인다.
    """

    def __init__(self, max_per_sec: float):
        self._interval = 1.0 / max_per_sec if max_per_sec > 0 else 0.0
        self._lock = threading.Lock()
        self._next = 0.0

    def wait(self) -> None:
        if self._interval <= 0:
            return
        with self._lock:
            now = time.monotonic()
            delay = self._next - now
            if delay > 0:
                time.sleep(delay)
                now = time.monotonic()
            self._next = max(now, self._next) + self._interval


# 캐시는 1차/2차 워커와 save_cache가 함께 만지므로 단일 락으로 보호한다.
# (dict 쓰기는 GIL로 원자적이지만 save_cache의 json.dumps가 순회 중이면 깨진다.)
_CACHE_LOCK = threading.Lock()
_RATE_LIMITER: RateLimiter | None = None


def cache_lock() -> threading.Lock:
    return _CACHE_LOCK


def set_rate_limiter(limiter: RateLimiter | None) -> None:
    """병렬 처리 동안만 전역 레이트리미터를 켠다. 직렬 처리(기본)에서는 None."""
    global _RATE_LIMITER
    _RATE_LIMITER = limiter


def load_cache(cache_file: Path) -> dict[str, Any]:
    if not cache_file.exists():
        return {}
    try:
        return json.loads(cache_file.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        # 이전 실행이 저장 도중 끊겨 캐시가 깨졌으면 버리고 새로 시작한다.
        return {}


def save_cache(cache_file: Path, cache: dict[str, Any]) -> None:
    # 병렬 워커가 cache를 쓰는 중에 직렬화하면 "dict changed size" 오류가 나므로
    # 캐시 락을 잡은 채로 스냅샷을 만든다.
    with _CACHE_LOCK:
        payload = json.dumps(cache, ensure_ascii=False, indent=2)
    tmp = cache_file.with_suffix(cache_file.suffix + ".tmp")
    tmp.write_text(payload, encoding="utf-8")
    tmp.replace(cache_file)


def juso_query(
    session: requests.Session,
    key: str,
    keyword: str,
    cache: dict[str, Any],
    count: int = 5,
    preserve_commas: bool = False,
) -> dict[str, Any]:
    # 모듈 상단에서 import하면 registry/__init__ → juso → juso_search 순환이 생기므로
    # 호출 시점에 가져온다 (sys.modules 캐시라 비용은 무시할 수준).
    from .registry.normalize import juso_keyword

    keyword = juso_keyword(keyword, preserve_commas=preserve_commas)
    if not keyword:
        return {"keyword": "", "total": 0, "rows": []}
    cache_key = f"{'raw' if preserve_commas else 'clean'}:{count}:{keyword}"
    with _CACHE_LOCK:
        if cache_key in cache:
            return cache[cache_key]
    limiter = _RATE_LIMITER
    if limiter is not None:
        limiter.wait()
    data = request_juso(key, keyword, count, timeout=15, session=session)
    if "error_code" in data:
        res = {
            "keyword": keyword,
            "total": 0,
            "rows": [],
            "error": data["error_message"],
        }
    else:
        res = {"keyword": keyword, "total": data["total"], "rows": data["rows"][:count]}
    with _CACHE_LOCK:
        cache[cache_key] = res
    if limiter is None:
        # 직렬 처리 기본 경로의 호출 간격 유지(레이트리미터가 켜지면 그쪽이 페이싱).
        time.sleep(0.04)
    return res
