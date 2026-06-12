from __future__ import annotations

from pathlib import Path
import time
from typing import Literal

import openpyxl

from .clients import JusoClient, KoreaPostRoadNameClient, SearchResult
from .normalizer import base_for_search, compact_for_epost, normalize_for_search


ProviderMode = Literal["none", "juso", "epost", "both"]
STATUS_NOT_FOUND = "검색주소없음"
STATUS_AMBIGUOUS = "2건이상검색"

# 로컬 정제 단계에서 검색 불가 판정된 사유의 사람용 설명 (검증상세 컬럼용)
LOCAL_STATUS_KO = {
    "invalid_marker": "원주소가 '미상' 등 무효 표기",
    "malformed": "주소 형식이 아님 (행정구역/번지 없음)",
    "unrecognized": "지번/도로명 골격을 찾지 못함",
}

# 연속 호출로 API가 차단되지 않도록 실제 호출 사이에만 두는 간격(초)
API_CALL_INTERVAL = 0.05


def col_to_index(col: str) -> int:
    col = col.strip().upper()
    value = 0
    for ch in col:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Invalid Excel column: {col}")
        value = value * 26 + ord(ch) - ord("A") + 1
    return value


def process_workbook(
    input_path: str | Path,
    output_path: str | Path,
    source_col: str = "H",
    target_col: str = "I",
    status_col: str | None = None,
    detail_col: str | None = None,
    provider: ProviderMode = "both",
    mark_missing: bool = False,
    header: bool = True,
) -> dict[str, int]:
    wb = openpyxl.load_workbook(input_path)
    ws = wb.active
    source_idx = col_to_index(source_col)
    target_idx = col_to_index(target_col)
    status_idx = col_to_index(status_col) if status_col else None
    detail_idx = col_to_index(detail_col) if detail_col else None
    if detail_idx and not status_idx:
        raise RuntimeError("--detail-col requires --status-col")

    if header:
        ws.cell(row=1, column=target_idx).value = "주소검색어"
        if status_idx:
            ws.cell(row=1, column=status_idx).value = "주소검색결과"
        if detail_idx:
            ws.cell(row=1, column=detail_idx).value = "주소검증상세"

    juso = JusoClient() if provider in ("juso", "both") else None
    epost = KoreaPostRoadNameClient() if provider in ("epost", "both") else None
    if juso is not None and not juso.key:
        juso = None
    if epost is not None and not epost.key:
        epost = None
    if mark_missing and provider != "none" and juso is None and epost is None:
        raise RuntimeError("At least one API key is required when --mark-missing validates provider results")

    stats = {
        "total": 0,
        "road": 0,
        "lot": 0,
        "empty": 0,
        "invalid": 0,
        "missing": 0,
        "ambiguous": 0,
        "verified": 0,
    }
    start_row = 2 if header else 1
    # 같은 원주소가 여러 행에 반복되는 파일이 흔해서 검증 결과를 재사용한다.
    verify_cache: dict[tuple[str, str], tuple[str, str]] = {}
    for row in range(start_row, ws.max_row + 1):
        raw = ws.cell(row=row, column=source_idx).value
        normalized = normalize_for_search(raw)
        ws.cell(row=row, column=target_idx).value = normalized.query
        stats["total"] += 1
        stats[normalized.kind if normalized.kind in stats else "invalid"] += 1

        if status_idx:
            status = ""
            verify_detail = ""
            if normalized.kind == "empty":
                # 원주소가 비어 있는 행(서식만 남은 말미 행 포함)은 검토 대상이
                # 아니므로 상태를 비워 둬 후단 자동화가 불량 주소와 혼동하지 않게 한다.
                pass
            elif not normalized.searchable:
                status = STATUS_NOT_FOUND
                verify_detail = LOCAL_STATUS_KO.get(normalized.status, normalized.status)
                stats["missing"] += 1
            elif mark_missing and (juso is not None or epost is not None):
                cache_key = (normalized.query, normalized.kind)
                cached = verify_cache.get(cache_key)
                if cached is None:
                    cached = _verify(normalized.query, normalized.kind, juso, epost)
                    verify_cache[cache_key] = cached
                verification, verify_detail = cached
                if verification == "verified":
                    stats["verified"] += 1
                elif verification == "ambiguous":
                    status = STATUS_AMBIGUOUS
                    stats["ambiguous"] += 1
                else:
                    status = STATUS_NOT_FOUND
                    stats["missing"] += 1
            ws.cell(row=row, column=status_idx).value = status
            if detail_idx:
                ws.cell(row=row, column=detail_idx).value = verify_detail

    wb.save(output_path)
    return stats


def _result_note(provider_label: str, result: SearchResult) -> str:
    if result.total_count == 1:
        road = result.first.get("roadAddr") or result.first.get("lnmAdres") or ""
        zip_no = result.first.get("zipNo") or result.first.get("zipNo1") or ""
        tail = f": {road}" + (f" (우){zip_no}" if zip_no else "") if road or zip_no else ""
        return f"{provider_label} 1건{tail}"
    return f"{provider_label} {result.total_count}건"


def _verify(query: str, kind: str, juso: JusoClient | None, epost: KoreaPostRoadNameClient | None) -> tuple[Literal["verified", "ambiguous", "missing"], str]:
    """(판정, 사람이 읽을 검증 상세) 반환.

    Juso는 상세 포함 검색이 0건이면 골격(시도~지번/건물번호)으로 한 번 더 검색한다.
    상세 표기('제비동 제101호' 등) 때문에 멀쩡한 주소가 불량 처리되는 것을 막고,
    골격조차 0건인 진짜 불량과 구분되도록 상세에 검색 경로를 남긴다.
    """
    if not query:
        return "missing", ""
    results: list[SearchResult] = []
    notes: list[str] = []
    if juso is not None:
        juso_queries = [("상세포함", query)]
        base = base_for_search(query, kind)
        if base and base != query:
            juso_queries.append(("골격", base))
        for label, juso_query in juso_queries:
            try:
                result = juso.search(juso_query, count=5)
            except Exception as exc:
                # 재시도 후에도 남은 전송 오류는 다른 provider 결과로 판정을 이어가고,
                # 모든 provider가 오류면 아래에서 실행을 중단한다.
                result = SearchResult("juso", 0, {"errorCode": "transport_error", "errorMessage": str(exc)}, "")
            time.sleep(API_CALL_INTERVAL)
            if result.has_error:
                notes.append(f"JUSO[{label}] 오류")
                break
            notes.append(_result_note(f"JUSO[{label}]", result))
            if result.total_count != 0:
                break
        results.append(result)
    if epost is not None:
        search_se = "road" if kind == "road" else "dong"
        epost_queries = [query]
        compact_query = compact_for_epost(query, kind)
        if compact_query and compact_query not in epost_queries:
            epost_queries.append(compact_query)
        for epost_query in epost_queries:
            try:
                result = epost.search(epost_query, search_se=search_se, count=5)
            except Exception as exc:
                result = SearchResult(
                    "epost",
                    0,
                    {"returnCode": "transport_error", "returnMessage": str(exc)},
                    "",
                )
            time.sleep(API_CALL_INTERVAL)
            results.append(result)
            if not result.has_error and result.total_count != 0:
                break
        notes.append("EPOST 오류" if result.has_error else _result_note("EPOST", result))
    usable_results = [result for result in results if not result.has_error]
    if not usable_results and results:
        raise RuntimeError("Address validation providers returned API errors; check API keys before marking missing addresses")
    detail = "; ".join(notes)
    if any(result.total_count >= 2 for result in usable_results):
        return "ambiguous", detail
    if any(result.total_count == 1 for result in usable_results):
        return "verified", detail
    return "missing", detail
