from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

from .clients import JusoClient, KoreaPostRoadNameClient
from .excel import collect_feedback, process_workbook
from .history import VerifyHistory
from .normalizer import compact_for_epost, normalize_for_search
from .regions import AdminDict
from .typo import load_typo_rules, set_extra_typo_rules


def main(argv: list[str] | None = None) -> int:
    cli_args = list(sys.argv[1:] if argv is None else argv)
    if cli_args and cli_args[0] == "registry":
        from .registry.cli import main as registry_main

        return registry_main(cli_args[1:])

    parser = argparse.ArgumentParser(prog="address-cleaner")
    sub = parser.add_subparsers(dest="command", required=True)

    # registry는 위에서 자체 파서로 위임되므로 여기 등록은 --help 노출용이다.
    sub.add_parser(
        "registry",
        help="법원 등기부등본 열람페이지 전체검색용 주소 생성 (registry-address-refine --help 참고)",
        add_help=False,
    )

    p_norm = sub.add_parser(
        "normalize", help="Normalize one address into an API-searchable query"
    )
    p_norm.add_argument("address")

    p_excel = sub.add_parser("excel", help="Normalize an Excel workbook")
    p_excel.add_argument("input")
    p_excel.add_argument("-o", "--output", required=True)
    p_excel.add_argument("--source-col", default="H")
    p_excel.add_argument("--target-col", default="I")
    p_excel.add_argument("--status-col")
    p_excel.add_argument(
        "--detail-col", help="검증 상세(검색 경로/건수/표준주소/우편번호)를 기록할 열"
    )
    p_excel.add_argument(
        "--provider", choices=["none", "juso", "epost", "both"], default="both"
    )
    p_excel.add_argument("--mark-missing", action="store_true")
    p_excel.add_argument(
        "--admin-dict",
        help="행안부 '법정동코드 전체자료' 텍스트 파일 경로 (API 호출 전 오프라인 행정구역 검증)",
    )
    p_excel.add_argument(
        "--history",
        help="검증 이력 SQLite 파일 경로 (최근 결과 재사용 + 판정 변화 감지)",
    )
    p_excel.add_argument(
        "--history-max-age-days",
        type=int,
        default=14,
        help="이력 재사용 허용 일수 (기본 14)",
    )
    p_excel.add_argument(
        "--corrections-out",
        help="교정 후보 리포트 JSON 저장 경로 (골격/지번변형으로만 통한 주소 쌍)",
    )
    p_excel.add_argument(
        "--typo-rules",
        type=Path,
        default=None,
        help='추가 오타 교정 규칙 JSON 경로 (예: [["프루지오", "푸르지오"]])',
    )
    p_excel.add_argument(
        "--workers",
        type=int,
        default=8,
        help="API 검증 병렬 워커 수 (기본 8, 1이면 직렬). 초당 호출은 자동으로 제한됩니다.",
    )

    p_feedback = sub.add_parser(
        "feedback", help="파워쉘 처리결과 엑셀의 실패 행을 모아 재정제 리포트 생성"
    )
    p_feedback.add_argument("input")
    p_feedback.add_argument(
        "-o", "--output", help="리포트 JSON 저장 경로 (생략 시 stdout)"
    )
    p_feedback.add_argument("--source-col", default="H")
    p_feedback.add_argument("--target-col", default="I")
    p_feedback.add_argument("--result-col", default="M")
    p_feedback.add_argument("--ps-detail-col", default="N")

    p_probe = sub.add_parser("probe", help="Probe configured API key with one query")
    p_probe.add_argument("provider", choices=["juso", "epost"])
    p_probe.add_argument("query")

    args = parser.parse_args(cli_args)
    if args.command == "normalize":
        normalized = normalize_for_search(args.address)
        print(json.dumps(normalized.__dict__, ensure_ascii=False, indent=2))
        return 0
    if args.command == "excel":
        history = None
        try:
            if args.typo_rules:
                set_extra_typo_rules(load_typo_rules(args.typo_rules))
            admin_dict = (
                AdminDict.load(Path(args.admin_dict)) if args.admin_dict else None
            )
            if args.history:
                history = VerifyHistory(
                    Path(args.history), max_age_days=args.history_max_age_days
                )
            stats = process_workbook(
                Path(args.input),
                Path(args.output),
                source_col=args.source_col,
                target_col=args.target_col,
                status_col=args.status_col,
                detail_col=args.detail_col,
                provider=args.provider,
                mark_missing=args.mark_missing,
                admin_dict=admin_dict,
                history=history,
                corrections_path=Path(args.corrections_out)
                if args.corrections_out
                else None,
                workers=args.workers,
            )
        except (RuntimeError, ValueError, OSError) as exc:
            print(f"error: {exc}", file=sys.stderr)
            return 2
        finally:
            if history is not None:
                history.close()
        print(json.dumps(stats, ensure_ascii=False, indent=2))
        return 0
    if args.command == "feedback":
        try:
            report = collect_feedback(
                Path(args.input),
                source_col=args.source_col,
                target_col=args.target_col,
                result_col=args.result_col,
                ps_detail_col=args.ps_detail_col,
            )
        except (RuntimeError, ValueError, OSError) as exc:
            print(f"error: {exc}", file=sys.stderr)
            return 2
        text = json.dumps(report, ensure_ascii=False, indent=2)
        if args.output:
            Path(args.output).write_text(text, encoding="utf-8")
            print(
                f"실패 {report['failures']}건 (재정제 변경 {report['requeryChanged']}건) → {args.output}"
            )
        else:
            print(text)
        return 0
    if args.command == "probe":
        if args.provider == "epost":
            epost_client = KoreaPostRoadNameClient()
            normalized = normalize_for_search(args.query)
            search_se = "road" if normalized.kind == "road" else "dong"
            queries = [normalized.query or args.query]
            compact_query = compact_for_epost(
                normalized.query or args.query, normalized.kind
            )
            if compact_query and compact_query not in queries:
                queries.append(compact_query)
            query_used = queries[0]
            result = epost_client.search(query_used, search_se=search_se)
            for query in queries[1:]:
                if not result.has_error and result.total_count != 0:
                    break
                result = epost_client.search(query, search_se=search_se)
                query_used = query
        else:
            result = JusoClient().search(args.query)
            query_used = args.query
        print(
            json.dumps(
                {
                    "provider": result.provider,
                    "query": query_used,
                    "found": result.found,
                    "total_count": result.total_count,
                    "first": result.first,
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
