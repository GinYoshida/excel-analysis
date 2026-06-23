"""Command-line entry point for xlsx-flow."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from xlsx_flow.parser.workbook import analyze
from xlsx_flow.render.html import render_html


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="xlsx-flow")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("analyze", help="Analyze an .xlsx and emit interactive HTML")
    p.add_argument("input", help="Path to the input .xlsx")
    p.add_argument("-o", "--output", default="out.html", help="Output HTML path")
    p.add_argument("--json", dest="json_out", default=None,
                   help="Also write the intermediate model JSON to this path")

    args = parser.parse_args(argv)

    if args.cmd == "analyze":
        model = analyze(args.input)
        html = render_html(model)
        Path(args.output).write_text(html, encoding="utf-8")
        if args.json_out:
            Path(args.json_out).write_text(model.to_json(), encoding="utf-8")
        print(f"wrote {args.output}"
              + (f" and {args.json_out}" if args.json_out else ""))
        if model.warnings:
            print(f"  ({len(model.warnings)} warning(s) — see the HTML panel)")
        return 0
    return 1


if __name__ == "__main__":
    sys.exit(main())
