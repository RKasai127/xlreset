"""Command-line entry point."""

import argparse
import sys
from pathlib import Path

from xlreset.workbook import XlresetError, reset_view

_EXIT_SUCCESS = 0
_EXIT_FAILURE = 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="xlreset",
        description=(
            "Reset every sheet in an Excel (.xlsx/.xlsm) file to cell A1 at 100% zoom, "
            "and make the first sheet the active tab."
        ),
    )
    parser.add_argument(
        "file",
        metavar="FILE",
        help="Path to a .xlsx/.xlsm file to rewrite in place.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    path = Path(args.file)
    try:
        reset_view(path)
    except (XlresetError, OSError) as exc:
        print(f"xlreset: {path}: {exc}", file=sys.stderr)
        return _EXIT_FAILURE
    return _EXIT_SUCCESS
