"""Command line entry point: ``docconvert INPUT OUTPUT`` or ``docconvert INPUT --to pdf``."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .core import SUPPORTED, convert


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="docconvert",
        description="Convert documents between " + ", ".join(SUPPORTED) + ".",
    )
    parser.add_argument("inputs", nargs="+", help="input file(s)")
    parser.add_argument("-o", "--output", help="output file (single input only)")
    parser.add_argument("-t", "--to", choices=SUPPORTED, help="target format (used when -o is omitted)")
    parser.add_argument("--no-soffice", action="store_true", help="never shell out to LibreOffice")
    args = parser.parse_args(argv)

    if args.output and len(args.inputs) > 1:
        parser.error("-o/--output only works with a single input; use --to for batches")
    if not args.output and not args.to:
        parser.error("give either -o OUTPUT or --to FORMAT")

    failed = 0
    for inp in args.inputs:
        src = Path(inp)
        dst = Path(args.output) if args.output else src.with_suffix("." + args.to)
        try:
            convert(src, dst, use_soffice=not args.no_soffice)
            print(f"{src} -> {dst}")
        except Exception as exc:  # noqa: BLE001 - report and continue with next file
            failed += 1
            print(f"error: {src}: {exc}", file=sys.stderr)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
