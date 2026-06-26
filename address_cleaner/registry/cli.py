"""CLI 진입점.

기존에 이 모듈에 있던 함수들은 normalize/juso/excel 모듈로 분리됐고,
하위 호환을 위해 여기서 그대로 re-export 한다.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from ..clients import juso_key_from_env
from .excel import (  # noqa: F401
    FINAL_COLUMNS,
    JUSO2_COLUMNS,
    JUSO_COLUMNS,
    REGISTRY_COLUMNS,
    WHOLE_COLUMNS,
    add_or_replace_sheet,
    ensure_columns,
    get_cell,
    refine,
    set_row_values,
)
from .juso import (  # noqa: F401
    API_URL,
    candidate_id,
    first_pass_status,
    juso_query,
    load_cache,
    make_queries,
    save_cache,
    score_candidate,
)
from .normalize import (  # noqa: F401
    KOR_DONG_MAP,
    LOT_RE,
    SIDO_RE,
    building_name,
    building_tokens,
    clean_raw,
    district_key,
    dong_key,
    find_lot,
    has_extra_parcels,
    is_probable_building_dong,
    juso_keyword,
    load_typo_rules,
    lot_key,
    lot_variants,
    norm,
    normalize_bld_dong,
    normalize_spaces,
    original_is_under_specified,
    parse_lot_addr,
    road_no_key,
    set_extra_typo_rules,
    strip_building_tail_after_lot,
    strip_detail,
    strip_unit,
    suffix_dong,
    suffix_ho,
    tail_after_actual_lot,
    tail_after_lot,
    typo_fix,
    typo_fix_first_pass,
    unit_extract,
    unit_out_of_range,
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="HUG 강제경매/우선신청 엑셀을 법원 부동산등기부등본 열람페이지 전체검색용 주소 파일로 정제합니다.")
    parser.add_argument("input", type=Path, help="입력 xlsx 파일")
    parser.add_argument("-o", "--output-dir", type=Path, default=Path("out"), help="출력 디렉터리 (기본: out)")
    parser.add_argument("--cache", type=Path, default=None, help="Juso API 캐시 JSON 경로")
    parser.add_argument("--typo-rules", type=Path, default=None, help='추가 오타 교정 규칙 JSON 경로 (예: [["프루지오", "푸르지오"]])')
    parser.add_argument("--workers", type=int, default=8, help="Juso 검색 병렬 워커 수 (기본 8, 1이면 직렬). 초당 호출은 자동으로 제한됩니다.")
    parser.add_argument("--quiet", action="store_true", help="진행 로그 최소화")
    parser.add_argument("--debug", action="store_true", help="오류 시 전체 traceback 출력")
    args = parser.parse_args(argv)

    key = juso_key_from_env()
    if not key:
        print("ERROR: JUSO_CONFM_KEY 또는 JUSO_API_KEY 환경변수가 필요합니다.", file=sys.stderr)
        return 2
    try:
        if args.typo_rules:
            set_extra_typo_rules(load_typo_rules(args.typo_rules))
        summary = refine(args.input, args.output_dir, key, args.cache, verbose=not args.quiet, workers=args.workers)
    except Exception as exc:
        if args.debug:
            raise
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
