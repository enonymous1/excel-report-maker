"""Command line: generate, validate, and list templates."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Sequence

from excel_report_maker import __version__
from excel_report_maker.config import list_templates, load_data, load_template
from excel_report_maker.engine import generate_report
from excel_report_maker.errors import DataError, TemplateError


def _add_template_arg(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--template",
        "-t",
        required=True,
        help="Template path, or a bundled template name (see list-templates).",
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-report-maker",
        description="Generate formatted Excel workbooks from JSON templates and data.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    generate = sub.add_parser("generate", help="Generate a workbook from a template and a data file.")
    _add_template_arg(generate)
    generate.add_argument("--data", "-d", required=True, help="JSON/JSON5 data file for the template.")
    generate.add_argument(
        "--output",
        "-o",
        help="Output path. Defaults to the template's output_dir/file_name.",
    )

    validate = sub.add_parser("validate", help="Check a template (and optionally data) without writing a file.")
    _add_template_arg(validate)
    validate.add_argument("--data", "-d", help="JSON/JSON5 data file to check against the template.")

    sub.add_parser("list-templates", help="List the templates available by name.")
    return parser


def _default_output_path(template: dict, output: str | None) -> Path:
    if output:
        return Path(output)
    output_dir = str(template.get("output_dir", "outputs"))
    file_name = str(template.get("file_name", "Generated_Report.xlsx"))
    return Path(output_dir) / file_name


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    if args.command == "list-templates":
        for name in list_templates():
            print(name)
        return 0

    try:
        template = load_template(args.template)
    except (FileNotFoundError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    if args.command == "validate":
        if args.data:
            try:
                load_data(args.data, template)
            except (ValueError, OSError) as exc:
                print(f"Invalid: {exc}", file=sys.stderr)
                return 1
        print(f"Valid: {args.template} ({template['report_type']}, {len(template['columns'])} columns)")
        return 0

    # generate
    try:
        data = load_data(args.data, template)
    except (ValueError, OSError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    output_path = _default_output_path(template, args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        generate_report(template, data, str(output_path))
    except (TemplateError, DataError, ValueError, OSError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    print(f"Generated: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
