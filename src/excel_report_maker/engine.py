"""Workbook generation: checklists and multi-section matrices via XlsxWriter.

The engine is dependency-light on purpose (XlsxWriter + stdlib only):
v1 routed checklist data through pandas solely to write it back out,
which is where the empty-data crash came from. Values are written
directly, so an empty checklist still produces a valid workbook with
just its header row, and a ragged row is a named error, not a
traceback from inside a DataFrame constructor.
"""

from __future__ import annotations

import os
from io import BytesIO
from pathlib import Path
from typing import Any, BinaryIO

import xlsxwriter
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet

from excel_report_maker.config import (
    strip_private_keys,
    validate_checklist_data,
    validate_matrix_data,
    validate_template,
)
from excel_report_maker.errors import DataError, TemplateError

# Named formats a template may reference without defining them, plus
# any name defined under the template's own ``styles`` registry.
BUILTIN_FORMATS: dict[str, dict[str, Any]] = {
    "green": {"bg_color": "#C6EFCE", "font_color": "#006100"},
    "yellow": {"bg_color": "#FFEB9C", "font_color": "#9C6500"},
    "red": {"bg_color": "#FFC7CE", "font_color": "#9C0006"},
    "alert_red": {"bg_color": "#FFC7CE", "font_color": "#9C0006"},
}

_FORMAT_CONTROL_KEYS = {"index", "name", "suffix", "width", "validation"}

_CONDITIONAL_FORMAT_KEYS = {"type", "criteria", "value", "minimum", "maximum", "stop_if_true"}

_MAX_AUTO_WIDTH = 100.0


def _format_props(raw: dict[str, Any], *, control_keys: set[str] | frozenset[str] | None = None) -> dict[str, Any]:
    """Format properties with documentation and control keys removed."""
    props = strip_private_keys(raw)
    excluded = control_keys or set()
    return {key: value for key, value in props.items() if key not in excluded}


def _resolve_format(
    workbook: Workbook, template: dict[str, Any], spec: Any, *, cache: dict[str, Any], source: str
) -> Any:
    """Turn a format spec (name, dict, or Format) into an XlsxWriter Format."""
    if spec is None:
        return None
    if isinstance(spec, str):
        if spec in cache:
            return cache[spec]
        if spec in BUILTIN_FORMATS:
            fmt = workbook.add_format(BUILTIN_FORMATS[spec])
        else:
            styles = template.get("styles")
            if not isinstance(styles, dict) or spec not in styles:
                known = sorted(set(BUILTIN_FORMATS) | set(styles if isinstance(styles, dict) else ()))
                raise TemplateError(f"{source}: unknown format name {spec!r}; known formats are {known}.")
            style = styles[spec]
            if not isinstance(style, dict):
                raise TemplateError(f"{source}: style {spec!r} must be a format object, not {type(style).__name__}.")
            fmt = workbook.add_format(_format_props(style))
        cache[spec] = fmt
        return fmt
    if isinstance(spec, dict):
        return workbook.add_format(_format_props(spec))
    # Already a Format object (programmatic callers).
    if hasattr(spec, "_get_dxf_index") or type(spec).__name__ == "Format":
        return spec
    raise TemplateError(f"{source}: format must be a name or a format object, got {spec!r}.")


def _clean_conditional_rule(
    workbook: Workbook, template: dict[str, Any], rule: Any, *, cache: dict[str, Any], source: str
) -> dict[str, Any]:
    """Validate one conditional-format rule and resolve its format."""
    if not isinstance(rule, dict):
        raise TemplateError(f"{source}: conditional format rules must be objects, got {type(rule).__name__}.")
    cleaned = strip_private_keys(rule)
    unknown = set(cleaned) - _CONDITIONAL_FORMAT_KEYS - {"format"}
    if unknown:
        raise TemplateError(f"{source}: unknown conditional-format option(s): {sorted(unknown)}.")
    if "format" not in cleaned:
        raise TemplateError(f"{source}: conditional format rule needs a 'format'.")
    result = {key: value for key, value in cleaned.items() if key != "format"}
    result["format"] = _resolve_format(workbook, template, cleaned["format"], cache=cache, source=source)
    return result


def _apply_page_setup(worksheet: Worksheet, template: dict[str, Any], *, source: str) -> None:
    """Apply a template's page_setup (orientation, paper size, margins)."""
    setup = template.get("page_setup")
    if setup is None:
        return
    if not isinstance(setup, dict):
        raise TemplateError(f"{source}: 'page_setup' must be an object.")
    orientation = setup.get("orientation")
    if orientation == "landscape":
        worksheet.set_landscape()
    elif orientation == "portrait":
        worksheet.set_portrait()
    elif orientation is not None:
        raise TemplateError(f"{source}: page_setup orientation must be 'landscape' or 'portrait'.")
    paper_size = setup.get("paper_size")
    if paper_size is not None:
        if not isinstance(paper_size, int) or isinstance(paper_size, bool):
            raise TemplateError(f"{source}: page_setup paper_size must be an integer code.")
        worksheet.set_paper(paper_size)
    margins = setup.get("margins")
    if margins is not None:
        if not isinstance(margins, dict):
            raise TemplateError(f"{source}: page_setup margins must be an object.")
        unknown_margins = set(margins) - {"left", "right", "top", "bottom"}
        if unknown_margins:
            raise TemplateError(
                f"{source}: page_setup margins only support left/right/top/bottom, got {sorted(unknown_margins)}."
            )
        # XlsxWriter's set_margins takes the four margins positionally.
        worksheet.set_margins(
            margins.get("left", 0.7),
            margins.get("right", 0.7),
            margins.get("top", 0.75),
            margins.get("bottom", 0.75),
        )


def _apply_freeze_panes(worksheet: Worksheet, template: dict[str, Any], *, source: str) -> None:
    freeze = template.get("freeze_panes")
    if freeze is None:
        return
    if (
        not isinstance(freeze, (list, tuple))
        or len(freeze) != 2
        or not all(isinstance(v, int) and not isinstance(v, bool) and v >= 0 for v in freeze)
    ):
        raise TemplateError(f"{source}: 'freeze_panes' must be [row, col] of non-negative integers.")
    worksheet.freeze_panes(freeze[0], freeze[1])


def _auto_width(values: list[Any], header: str) -> float:
    """Width from the longest cell, as in v1, capped so one long note
    does not blow the sheet out. Empty/blank cells do not count."""
    longest = 0
    for value in values:
        if value is None:
            continue
        text = str(value)
        if text:
            longest = max(longest, len(text))
    return min(max(longest, len(header) * 1.2) + 4, _MAX_AUTO_WIDTH)


def _column_indices(spec: Any, *, column_count: int, source: str) -> list[int]:
    indices = [spec] if isinstance(spec, int) and not isinstance(spec, bool) else spec
    if not isinstance(indices, list) or not all(isinstance(i, int) and not isinstance(i, bool) for i in indices):
        raise TemplateError(f"{source}: column 'index' must be an integer or a list of integers.")
    for index in indices:
        if not 0 <= index < column_count:
            raise TemplateError(f"{source}: column index {index} is out of range for {column_count} columns.")
    return list(indices)


def _default_output(template: dict[str, Any]) -> str:
    output_dir = template.get("output_dir", "outputs")
    file_name = template.get("file_name", "report.xlsx")
    path = Path(output_dir) / file_name
    path.parent.mkdir(parents=True, exist_ok=True)
    return str(path)


def generate_checklist(
    template_config: dict[str, Any],
    data_payload: Any,
    buffer: str | Path | BinaryIO | None = None,
) -> Any:
    """Write a checklist workbook; returns the path written, or *buffer*.

    ``buffer`` may be a binary stream (the v1 call shape used by the
    Streamlit UI) or an explicit output path. Without it, the file
    goes to the template's ``output_dir``/``file_name`` (default
    ``outputs/report.xlsx``).
    """
    template = validate_template(template_config)
    columns: list[str] = template["columns"]
    rows = validate_checklist_data(data_payload, column_count=len(columns))
    sheet_name: str = template["sheet_name"]
    cache: dict[str, Any] = {}

    target: Any = buffer if buffer is not None else _default_output(template)
    if isinstance(target, (str, Path)) and not isinstance(target, BytesIO):
        Path(target).parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(target)
    try:
        worksheet = workbook.add_worksheet(sheet_name)
        header_fmt = workbook.add_format(_format_props(template.get("header_format", {})))
        _apply_freeze_panes(worksheet, template, source="checklist template")
        _apply_page_setup(worksheet, template, source="checklist template")

        # Suffixes (e.g. a trailing ":") apply to text only, and are
        # never doubled if the value already carries one.
        suffixes: dict[int, str] = {}
        for col_fmt in template.get("column_formats", []):
            if not isinstance(col_fmt, dict):
                raise TemplateError("checklist template: 'column_formats' entries must be objects.")
            for index in _column_indices(col_fmt.get("index"), column_count=len(columns), source="checklist template"):
                suffix = col_fmt.get("suffix")
                if isinstance(suffix, str) and suffix:
                    suffixes[index] = suffix

        processed: list[list[Any]] = []
        for row in rows:
            new_row = list(row)
            for index, suffix in suffixes.items():
                value = new_row[index]
                if isinstance(value, str) and value and not value.endswith(suffix):
                    new_row[index] = f"{value}{suffix}"
            processed.append(new_row)

        worksheet.write_row(0, 0, columns, header_fmt)
        for row_index, row in enumerate(processed, start=1):
            worksheet.write_row(row_index, 0, row)

        if processed:
            worksheet.add_table(
                0,
                0,
                len(processed),
                len(columns) - 1,
                {
                    "columns": [{"header": col} for col in columns],
                    "style": template.get("table_style", "Table Style Light 9"),
                },
            )
            # add_table rewrites the header row; restore the template style.
            worksheet.set_row(0, template.get("header_row_height", 20))
            for col_num, value in enumerate(columns):
                worksheet.write(0, col_num, value, header_fmt)
        else:
            worksheet.set_row(0, template.get("header_row_height", 20))

        covered: set[int] = set()
        for col_fmt in template.get("column_formats", []):
            indices = _column_indices(col_fmt.get("index"), column_count=len(columns), source="checklist template")
            covered.update(indices)
            cell_fmt = workbook.add_format(_format_props(col_fmt, control_keys=_FORMAT_CONTROL_KEYS))
            width = col_fmt.get("width")
            if width is not None and (not isinstance(width, (int, float)) or isinstance(width, bool) or width <= 0):
                raise TemplateError("checklist template: column 'width' must be a positive number.")
            for index in indices:
                column_values = [row[index] for row in processed]
                column_width = float(width) if width is not None else _auto_width(column_values, columns[index])
                worksheet.set_column(index, index, column_width, cell_fmt)
                validation = col_fmt.get("validation")
                if validation is not None and processed:
                    if not isinstance(validation, dict):
                        raise TemplateError("checklist template: column 'validation' must be an object.")
                    worksheet.data_validation(1, index, len(processed), index, strip_private_keys(validation))

        # Columns with no column_formats entry still get an auto-fit width.
        for index, header in enumerate(columns):
            if index not in covered:
                worksheet.set_column(index, index, _auto_width([row[index] for row in processed], header))

        for group in template.get("conditional_formats", []):
            if not isinstance(group, dict):
                raise TemplateError("checklist template: 'conditional_formats' entries must be objects.")
            indices = _column_indices(group.get("index"), column_count=len(columns), source="checklist template")
            rules = group.get("rules")
            if not isinstance(rules, list):
                raise TemplateError("checklist template: conditional format groups need a 'rules' list.")
            if not processed:
                continue
            for index in indices:
                for rule in rules:
                    options = _clean_conditional_rule(
                        workbook, template, rule, cache=cache, source="checklist template"
                    )
                    worksheet.conditional_format(1, index, len(processed), index, options)
    except Exception:
        try:
            workbook.close()
        except Exception:
            pass
        raise
    workbook.close()
    return target


def generate_matrix(
    template_config: dict[str, Any],
    data_payload: Any,
    buffer: str | Path | BinaryIO | None = None,
) -> Any:
    """Write a multi-section matrix workbook; returns the path, or *buffer*.

    ``buffer`` may be a binary stream or an explicit output path.

    Which column is highlighted as the score column comes from the
    template's ``score_column`` (v1 hardcoded column 5), and the
    width cap from ``width_cap`` (v1 hardcoded column 6 at 100).
    """
    template = validate_template(template_config)
    columns: list[str] = template["columns"]
    sections = validate_matrix_data(data_payload, column_count=len(columns))
    sheet_name: str = template["sheet_name"]
    score_column = template.get("score_column")

    target: Any = buffer if buffer is not None else _default_output(template)
    if isinstance(target, (str, Path)) and not isinstance(target, BytesIO):
        Path(target).parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(target)
    try:
        worksheet = workbook.add_worksheet(sheet_name)
        header_fmt = workbook.add_format(_format_props(template.get("header_format", {})))
        section_fmt = workbook.add_format(_format_props(template.get("section_header_format", {})))
        cell_fmt = workbook.add_format(_format_props(template.get("cell_format", {})))
        score_fmt = workbook.add_format(_format_props(template.get("score_highlight_format", {})))
        _apply_freeze_panes(worksheet, template, source="matrix template")
        _apply_page_setup(worksheet, template, source="matrix template")

        worksheet.write_row(0, 0, columns, header_fmt)
        worksheet.set_row(0, template.get("header_row_height", 20))

        max_lengths = [len(col) for col in columns]
        current_row = 1
        for section in sections:
            worksheet.merge_range(current_row, 0, current_row, len(columns) - 1, section["header"], section_fmt)
            current_row += 1
            for row in section["rows"]:
                for col_num, value in enumerate(row):
                    fmt = score_fmt if score_column is not None and col_num == score_column else cell_fmt
                    worksheet.write(current_row, col_num, value, fmt)
                    if value is not None and str(value):
                        max_lengths[col_num] = max(max_lengths[col_num], len(str(value)))
                current_row += 1

        width_cap = template.get("width_cap", _MAX_AUTO_WIDTH)
        if not isinstance(width_cap, (int, float)) or isinstance(width_cap, bool) or width_cap <= 0:
            raise TemplateError("matrix template: 'width_cap' must be a positive number.")
        explicit_widths = template.get("column_widths")
        if explicit_widths is not None:
            if not isinstance(explicit_widths, list) or len(explicit_widths) != len(columns):
                raise TemplateError("matrix template: 'column_widths' must list one width per column.")
        for index, length in enumerate(max_lengths):
            if explicit_widths is not None and explicit_widths[index] is not None:
                width = float(explicit_widths[index])
            else:
                width = min(length + 4, float(width_cap))
            worksheet.set_column(index, index, width)
    except Exception:
        try:
            workbook.close()
        except Exception:
            pass
        raise
    workbook.close()
    return target


_GENERATORS = {"checklist": generate_checklist, "matrix": generate_matrix}


def generate_report(
    template_config: dict[str, Any], data_payload: Any, buffer: str | Path | BinaryIO | None = None
) -> Any:
    """Dispatch to the generator for the template's ``report_type``."""
    template = validate_template(template_config)
    generator = _GENERATORS[template["report_type"]]
    return generator(template, data_payload, buffer)


def main() -> int:
    """Batch-generate every config in ``configs/`` and ``local_configs/``.

    Kept for v1 compatibility. Configs in ``configs/templates/`` are
    templates without data and are skipped; only configs carrying
    their own ``data``/``sections`` payload produce a file.
    """
    config_dirs = ["configs", "local_configs"]
    config_files: list[str] = []
    for config_dir in config_dirs:
        if os.path.isdir(config_dir):
            config_files.extend(sorted(str(path) for path in Path(config_dir).glob("*.json")))
    if not config_files:
        print(f"No configuration files found in {config_dirs}.")
        return 1

    from excel_report_maker.config import load_json_file

    failures = 0
    for config_file in config_files:
        try:
            config = load_json_file(config_file)
            if not isinstance(config, dict):
                print(f"Skipping {config_file}: not a JSON object.")
                continue
            report_type = config.get("report_type")
            if report_type not in _GENERATORS:
                print(f"Unknown report type {report_type!r} in {config_file}")
                failures += 1
                continue
            key = "data" if report_type == "checklist" else "sections"
            if key not in config:
                print(f"Skipping {config_file}: template without an embedded {key!r} payload.")
                continue
            result = _GENERATORS[report_type](config, config[key])
            print(f"Generated {report_type}: {result}")
        except (TemplateError, DataError, ValueError, OSError) as exc:
            print(f"Error processing {config_file}: {exc}")
            failures += 1
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
