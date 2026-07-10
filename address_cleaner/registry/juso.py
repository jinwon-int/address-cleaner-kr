"""Juso(도로명주소) 후보 검색어 생성·스코어링 (등기 모드 전용).

캐시·레이트리미터·juso_query 등 검색 인프라는 일반 excel 모드와 공유하는
address_cleaner/juso_search.py로 옮겨졌고, 기존 import 경로 하위호환을 위해
여기서 그대로 re-export 한다.
"""

from __future__ import annotations

from typing import Any

from ..clients import JUSO_ENDPOINT
from ..juso_search import (  # noqa: F401
    DEFAULT_CACHE_MAX_AGE_DAYS,
    RateLimiter,
    cache_entry_fresh,
    cache_lock,
    juso_query,
    load_cache,
    save_cache,
    set_cache_max_age_days,
    set_rate_limiter,
    stamp_cache_entry,
)
from .normalize import (
    building_tokens,
    clean_raw,
    district_key,
    dong_key,
    lot_key,
    lot_variants,
    norm,
    road_no_key,
    strip_building_tail_after_lot,
    strip_unit,
)

API_URL = JUSO_ENDPOINT


def first_pass_status(
    full: dict[str, Any], normalized: dict[str, Any] | None
) -> tuple[str, dict[str, Any]]:
    if full["total"] == 1:
        return "검색가능_단일_원문", full
    if full["total"] >= 2:
        return "다중검출_원문", full
    if normalized and normalized["total"] == 1:
        return "검색가능_단일_상세주소제거", normalized
    if normalized and normalized["total"] >= 2:
        return "다중검출_상세주소제거", normalized
    return "검색불가", full


def make_queries(
    raw: Any, final: Any, road_complete: Any, jibun_complete: Any
) -> list[str]:
    bases: list[str] = []
    for x in [raw, final, road_complete, jibun_complete]:
        c = clean_raw(x)
        if c:
            bases.extend([c, strip_unit(c), strip_building_tail_after_lot(c)])
            bases.extend(lot_variants(c))
    all_cleaned = [
        clean_raw(x)
        for x in [raw, final, road_complete, jibun_complete]
        if clean_raw(x)
    ]
    for c in all_cleaned:
        rk = road_no_key(c)
        lk = lot_key(c)
        dk = district_key(c)
        if rk:
            bases.append(norm(f"{dk} {rk}" if dk else rk))
        if lk:
            bases.append(norm(f"{dk} {lk}" if dk else lk))
    # If a road/lot number is malformed but a building name is distinctive,
    # Juso often resolves with district + building name (e.g. 김포시 테라스테이).
    districts = [district_key(c) for c in all_cleaned if district_key(c)]
    for dk in districts[:2]:
        for token in building_tokens(raw, final, road_complete, jibun_complete)[:4]:
            bases.append(norm(f"{dk} {token}"))
    seen: set[str] = set()
    out: list[str] = []
    for q in bases:
        q = norm(q)
        if len(q) >= 6 and q not in seen:
            seen.add(q)
            out.append(q)
    return out[:12]


def candidate_id(row: dict[str, Any]) -> str:
    return "|".join(
        str(row.get(k, ""))
        for k in [
            "admCd",
            "rnMgtSn",
            "udrtYn",
            "buldMnnm",
            "buldSlno",
            "bdMgtSn",
            "lnbrMnnm",
            "lnbrSlno",
        ]
    )


def score_candidate(
    row: dict[str, Any],
    raw: Any,
    final: Any,
    road_complete: Any,
    jibun_complete: Any,
    hit_queries: list[dict[str, Any]],
) -> tuple[int, list[str]]:
    combined = norm(
        " ".join(
            str(row.get(k, "") or "")
            for k in ["roadAddr", "roadAddrPart1", "jibunAddr", "bdNm"]
        )
    )
    road = norm(row.get("roadAddr") or row.get("roadAddrPart1") or "")
    jibun = norm(row.get("jibunAddr") or "")
    texts = [
        clean_raw(raw),
        clean_raw(final),
        clean_raw(road_complete),
        clean_raw(jibun_complete),
    ]
    score = 0
    reasons: list[str] = []
    for lk in {lot_key(t) for t in texts if lot_key(t)}:
        if lk and lk in jibun:
            score += 60
            reasons.append(f"지번일치:{lk}")
            break
    for rk in {road_no_key(t) for t in texts if road_no_key(t)}:
        if rk and rk in road:
            score += 55
            reasons.append(f"도로명건물번호일치:{rk}")
            break
    for dk in {district_key(t) for t in texts if district_key(t)}:
        if dk and dk in combined:
            score += 10
            reasons.append(f"시군구일치:{dk}")
            break
    for dg in {dong_key(t) for t in texts if dong_key(t)}:
        if dg and dg in jibun:
            score += 10
            reasons.append(f"법정동일치:{dg}")
            break
    matched = [
        t
        for t in building_tokens(raw, final, road_complete, jibun_complete)
        if t in combined.lower()
    ]
    if matched:
        score += min(25, 8 * len(matched))
        reasons.append("건물명일치:" + ",".join(matched[:3]))
    if len(hit_queries) == 1:
        score += 3
        reasons.append("단일검색경로")
    elif len(hit_queries) >= 2:
        score += 8
        reasons.append(f"복수검색경로:{len(hit_queries)}")
    return score, reasons
