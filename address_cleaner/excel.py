from __future__ import annotations

from pathlib import Path
from typing import Literal

import openpyxl

from .clients import JusoClient, KoreaPostRoadNameClient
from .normalizer import normalize_for_search


ProviderMode = Literal["none", "juso", "epost", "both"]


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
    provider: ProviderMode = "none",
    mark_missing: bool = False,
    header: bool = True,
) -> dict[str, int]:
    wb = openpyxl.load_workbook(input_path)
    ws = wb.active
    source_idx = col_to_index(source_col)
    target_idx = col_to_index(target_col)
    status_idx = col_to_index(status_col) if status_col else None

    if header:
        ws.cell(row=1, column=target_idx).value = "우체국API검색어"
        if status_idx:
            ws.cell(row=1, column=status_idx).value = "주소검색결과"

    juso = JusoClient() if provider in ("juso", "both") else None
    epost = KoreaPostRoadNameClient() if provider in ("epost", "both") else None

    stats = {"total": 0, "road": 0, "lot": 0, "fallback": 0, "missing": 0, "verified": 0}
    start_row = 2 if header else 1
    for row in range(start_row, ws.max_row + 1):
        raw = ws.cell(row=row, column=source_idx).value
        normalized = normalize_for_search(raw)
        ws.cell(row=row, column=target_idx).value = normalized.query
        stats["total"] += 1
        stats[normalized.kind if normalized.kind in stats else "fallback"] += 1

        if status_idx:
            status = ""
            if mark_missing:
                found = _verify(normalized.query, normalized.kind, juso, epost)
                if found:
                    stats["verified"] += 1
                else:
                    status = "주소없음"
                    stats["missing"] += 1
            elif not normalized.searchable:
                status = "주소없음"
                stats["missing"] += 1
            ws.cell(row=row, column=status_idx).value = status

    wb.save(output_path)
    return stats


def _verify(query: str, kind: str, juso: JusoClient | None, epost: KoreaPostRoadNameClient | None) -> bool:
    if not query:
        return False
    if juso is not None:
        result = juso.search(query, count=1)
        if result.found:
            return True
    if epost is not None:
        search_se = "road" if kind == "road" else "dong"
        result = epost.search(query, search_se=search_se, count=1)
        if result.found:
            return True
    return False

