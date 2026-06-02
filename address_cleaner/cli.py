from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

from .clients import JusoClient, KoreaPostRoadNameClient
from .excel import process_workbook
from .normalizer import normalize_for_search


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="address-cleaner")
    sub = parser.add_subparsers(dest="command", required=True)

    p_norm = sub.add_parser("normalize", help="Normalize one address into an API-searchable query")
    p_norm.add_argument("address")

    p_excel = sub.add_parser("excel", help="Normalize an Excel workbook")
    p_excel.add_argument("input")
    p_excel.add_argument("-o", "--output", required=True)
    p_excel.add_argument("--source-col", default="H")
    p_excel.add_argument("--target-col", default="I")
    p_excel.add_argument("--status-col")
    p_excel.add_argument("--provider", choices=["none", "juso", "epost", "both"], default="both")
    p_excel.add_argument("--mark-missing", action="store_true")

    p_probe = sub.add_parser("probe", help="Probe configured API key with one query")
    p_probe.add_argument("provider", choices=["juso", "epost"])
    p_probe.add_argument("query")

    args = parser.parse_args(argv)
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
            kind = normalize_for_search(args.query).kind
            result = client.search(args.query, search_se="road" if kind == "road" else "dong")
        else:
            result = client.search(args.query)
        print(json.dumps({"provider": result.provider, "found": result.found, "total_count": result.total_count, "first": result.first}, ensure_ascii=False, indent=2))
        return 0
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
