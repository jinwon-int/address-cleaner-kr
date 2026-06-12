from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

from .clients import JusoClient, KoreaPostRoadNameClient
from .excel import process_workbook
from .normalizer import compact_for_epost, normalize_for_search


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

    p_norm = sub.add_parser("normalize", help="Normalize one address into an API-searchable query")
    p_norm.add_argument("address")

    p_excel = sub.add_parser("excel", help="Normalize an Excel workbook")
    p_excel.add_argument("input")
    p_excel.add_argument("-o", "--output", required=True)
    p_excel.add_argument("--source-col", default="H")
    p_excel.add_argument("--target-col", default="I")
    p_excel.add_argument("--status-col")
    p_excel.add_argument("--detail-col", help="검증 상세(검색 경로/건수/표준주소/우편번호)를 기록할 열")
    p_excel.add_argument("--provider", choices=["none", "juso", "epost", "both"], default="both")
    p_excel.add_argument("--mark-missing", action="store_true")

    p_probe = sub.add_parser("probe", help="Probe configured API key with one query")
    p_probe.add_argument("provider", choices=["juso", "epost"])
    p_probe.add_argument("query")

    args = parser.parse_args(cli_args)
    if args.command == "normalize":
        normalized = normalize_for_search(args.address)
        print(json.dumps(normalized.__dict__, ensure_ascii=False, indent=2))
        return 0
    if args.command == "excel":
        try:
            stats = process_workbook(
                Path(args.input),
                Path(args.output),
                source_col=args.source_col,
                target_col=args.target_col,
                status_col=args.status_col,
                detail_col=args.detail_col,
                provider=args.provider,
                mark_missing=args.mark_missing,
            )
        except RuntimeError as exc:
            print(f"error: {exc}", file=sys.stderr)
            return 2
        print(json.dumps(stats, ensure_ascii=False, indent=2))
        return 0
    if args.command == "probe":
        client = JusoClient() if args.provider == "juso" else KoreaPostRoadNameClient()
        if args.provider == "epost":
            normalized = normalize_for_search(args.query)
            search_se = "road" if normalized.kind == "road" else "dong"
            queries = [normalized.query or args.query]
            compact_query = compact_for_epost(normalized.query or args.query, normalized.kind)
            if compact_query and compact_query not in queries:
                queries.append(compact_query)
            result = None
            query_used = queries[0]
            for query in queries:
                result = client.search(query, search_se=search_se)
                query_used = query
                if not result.has_error and result.total_count != 0:
                    break
        else:
            result = client.search(args.query)
            query_used = args.query
        print(json.dumps({"provider": result.provider, "query": query_used, "found": result.found, "total_count": result.total_count, "first": result.first}, ensure_ascii=False, indent=2))
        return 0
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
