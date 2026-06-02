from __future__ import annotations

from pathlib import Path
from typing import Literal

import openpyxl

from .clients import JusoClient, KoreaPostRoadNameClient, SearchResult
from .normalizer import normalize_for_search


ProviderMode = Literal["none", "juso", "epost", "both"]
STATUS_NOT_FOUND = "검색주소없음"
STATUS_AMBIGUOUS = "2건이상검색"


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
    provider: ProviderMode = "both",
    mark_missing: bool = False,
    header: bool = True,
) -> dict[str, int]:
    wb = openpyxl.load_workbook(input_path)
    ws = wb.active
    source_idx = col_to_index(source_col)
    target_idx = col_to_index(target_col)
    status_idx = col_to_index(status_col) if status_col else None

    if header:
        ws.cell(row=1, column=target_idx).value = "주소검색어"
        if status_idx:
            ws.cell(row=1, column=status_idx).value = "주소검색결과"

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
        "invalid": 0,
        "missing": 0,
        "ambiguous": 0,
        "verified": 0,
    }
    start_row = 2 if header else 1
    for row in range(start_row, ws.max_row + 1):
        raw = ws.cell(row=row, column=source_idx).value
        normalized = normalize_for_search(raw)
        ws.cell(row=row, column=target_idx).value = normalized.query
        stats["total"] += 1
        stats[normalized.kind if normalized.kind in stats else "invalid"] += 1

        if status_idx:
            status = ""
            if not normalized.searchable:
                status = STATUS_NOT_FOUND
                stats["missing"] += 1
            elif mark_missing and (juso is not None or epost is not None):
                verification = _verify(normalized.query, normalized.kind, juso, epost)
                if verification == "verified":
                    stats["verified"] += 1
                elif verification == "ambiguous":
                    status = STATUS_AMBIGUOUS
                    stats["ambiguous"] += 1
                else:
                    status = STATUS_NOT_FOUND
                    stats["missing"] += 1
            ws.cell(row=row, column=status_idx).value = status

    wb.save(output_path)
    return stats


def _verify(query: str, kind: str, juso: JusoClient | None, epost: KoreaPostRoadNameClient | None) -> Literal["verified", "ambiguous", "missing"]:
    if not query:
        return "missing"
    results: list[SearchResult] = []
    if juso is not None:
        results.append(juso.search(query, count=5))
    if epost is not None:
        search_se = "road" if kind == "road" else "dong"
        results.append(epost.search(query, search_se=search_se, count=5))
    usable_results = [result for result in results if not result.has_error]
    if not usable_results and results:
        raise RuntimeError("Address validation providers returned API errors; check API keys before marking missing addresses")
    if any(result.total_count >= 2 for result in usable_results):
        return "ambiguous"
    if any(result.total_count == 1 for result in usable_results):
        return "verified"
    return "missing"
