"""Loading and validation of JSON/JSON5 templates and data payloads.

Templates and data are plain JSON, with JSON5 comments allowed so a
template can carry notes for the next person editing it. Keys starting
with ``_`` inside format/rule dictionaries are documentation only and
are stripped before anything is handed to XlsxWriter: in v1 they were
passed through verbatim, which silently dropped every conditional
format in the master reference config.
"""

from __future__ import annotations

import importlib.resources
import json
from pathlib import Path
from typing import Any

import json5

from excel_report_maker.errors import DataError, TemplateError

REPORT_TYPES = ("checklist", "matrix")

# Where a template name is looked up when it is not an existing path:
# the templates bundled with the package, then the editable copies in
# the repository checkout.
_TEMPLATE_SEARCH_DIRS: tuple[Path, ...] = ()


def _bundled_templates_dir() -> Path:
    return Path(str(importlib.resources.files("excel_report_maker") / "templates"))


def load_json_file(path: str | Path) -> Any:
    """Parse a JSON/JSON5 file, naming the file in any parse error."""
    path = Path(path)
    try:
        with path.open("r", encoding="utf-8") as handle:
            return json5.load(handle)
    except (OSError, ValueError) as exc:
        raise ValueError(f"Could not parse {path}: {exc}") from exc


def parse_json_text(text: str, *, source: str = "<input>") -> Any:
    """Parse JSON/JSON5 text, naming the source in any parse error."""
    try:
        return json5.loads(text)
    except ValueError as exc:
        raise ValueError(f"Could not parse JSON from {source}: {exc}") from exc


def strip_private_keys(mapping: dict[str, Any]) -> dict[str, Any]:
    """Return a copy of *mapping* without documentation-only ``_`` keys."""
    return {key: value for key, value in mapping.items() if not key.startswith("_")}


def _search_dirs() -> list[Path]:
    dirs = [_bundled_templates_dir(), Path.cwd() / "configs" / "templates"]
    return [d for d in dirs if d.is_dir()]


def resolve_template_path(name_or_path: str) -> Path:
    """Resolve a template argument to a file: path first, then bundled name."""
    candidate = Path(name_or_path)
    if candidate.is_file():
        return candidate
    for directory in _search_dirs():
        for candidate_name in (name_or_path, f"{name_or_path}.json"):
            found = directory / candidate_name
            if found.is_file():
                return found
    searched = ", ".join(str(d) for d in _search_dirs())
    raise FileNotFoundError(
        f"Template {name_or_path!r} not found as a file or in the searched template directories ({searched})."
    )


def list_templates() -> list[str]:
    """Names of the templates available by name (bundled + checkout)."""
    names: dict[str, Path] = {}
    for directory in _search_dirs():
        for path in sorted(directory.glob("*.json")):
            names.setdefault(path.stem, path)
    return sorted(names)


def load_template(name_or_path: str | Path) -> dict[str, Any]:
    """Load and validate a template, returning its parsed dict."""
    path = Path(name_or_path) if isinstance(name_or_path, Path) else resolve_template_path(name_or_path)
    parsed = load_json_file(path)
    if not isinstance(parsed, dict):
        raise TemplateError(f"Template {path} must contain a JSON object, not {type(parsed).__name__}.")
    return validate_template(parsed, source=str(path))


def validate_template(template: dict[str, Any], *, source: str = "template") -> dict[str, Any]:
    """Fail loudly on a template that cannot produce a workbook."""
    report_type = template.get("report_type")
    if report_type not in REPORT_TYPES:
        raise TemplateError(f"{source}: 'report_type' must be one of {REPORT_TYPES}, got {report_type!r}.")
    columns = template.get("columns")
    if not isinstance(columns, list) or not columns or not all(isinstance(col, str) and col.strip() for col in columns):
        raise TemplateError(f"{source}: 'columns' must be a non-empty list of header strings.")
    sheet_name = template.get("sheet_name")
    if not isinstance(sheet_name, str) or not sheet_name.strip():
        raise TemplateError(f"{source}: 'sheet_name' must be a non-empty string.")
    if len(sheet_name) > 31:
        raise TemplateError(f"{source}: sheet name {sheet_name!r} exceeds Excel's 31-character limit.")
    file_name = template.get("file_name")
    if file_name is not None and (
        not isinstance(file_name, str)
        or not file_name.strip()
        or "/" in file_name
        or "\\" in file_name
        or not file_name.lower().endswith(".xlsx")
    ):
        raise TemplateError(f"{source}: 'file_name' must be a bare filename ending in .xlsx, got {file_name!r}.")
    if report_type == "matrix":
        score_column = template.get("score_column")
        if score_column is not None and (
            not isinstance(score_column, int) or isinstance(score_column, bool) or not 0 <= score_column < len(columns)
        ):
            raise TemplateError(f"{source}: 'score_column' must be a column index in range, got {score_column!r}.")
    for key in ("column_formats", "conditional_formats"):
        value = template.get(key)
        if value is not None and not isinstance(value, list):
            raise TemplateError(f"{source}: {key!r} must be a list, got {type(value).__name__}.")
    return template


def validate_checklist_data(data: Any, *, column_count: int, source: str = "data") -> list[list[Any]]:
    """Check checklist rows: a list of rows, each with one value per column."""
    if not isinstance(data, list):
        raise DataError(f"{source}: checklist data must be a list of rows, got {type(data).__name__}.")
    rows: list[list[Any]] = []
    for index, row in enumerate(data):
        if not isinstance(row, (list, tuple)):
            raise DataError(f"{source}: row {index} must be a list, got {type(row).__name__}.")
        if len(row) != column_count:
            raise DataError(f"{source}: row {index} has {len(row)} value(s), expected {column_count} (one per column).")
        rows.append(list(row))
    return rows


def validate_matrix_data(data: Any, *, column_count: int, source: str = "data") -> list[dict[str, Any]]:
    """Check matrix sections: a list of {header, rows} dicts, rows sized to columns."""
    if not isinstance(data, list):
        raise DataError(f"{source}: matrix data must be a list of sections, got {type(data).__name__}.")
    sections: list[dict[str, Any]] = []
    for index, section in enumerate(data):
        if not isinstance(section, dict):
            raise DataError(f"{source}: section {index} must be an object with 'header' and 'rows'.")
        header = section.get("header")
        if not isinstance(header, str) or not header.strip():
            raise DataError(f"{source}: section {index} needs a non-empty string 'header'.")
        raw_rows = section.get("rows")
        if not isinstance(raw_rows, list):
            raise DataError(f"{source}: section {index} ('{header}') needs 'rows' as a list.")
        rows: list[list[Any]] = []
        for row_index, row in enumerate(raw_rows):
            if not isinstance(row, (list, tuple)):
                raise DataError(f"{source}: section {index} row {row_index} must be a list, got {type(row).__name__}.")
            if len(row) != column_count:
                raise DataError(
                    f"{source}: section {index} row {row_index} has {len(row)} value(s), "
                    f"expected {column_count} (one per column)."
                )
            rows.append(list(row))
        sections.append({"header": header, "rows": rows})
    return sections


def load_data(path: str | Path, template: dict[str, Any]) -> list[Any]:
    """Load a data file and validate it against *template*."""
    parsed = load_json_file(path)
    columns = template["columns"]
    if template.get("report_type") == "matrix":
        return validate_matrix_data(parsed, column_count=len(columns), source=str(path))
    return validate_checklist_data(parsed, column_count=len(columns), source=str(path))


def dump_json(data: Any) -> str:
    """Serialize template/data dicts as plain JSON (comments are not preserved)."""
    return json.dumps(data, indent=4, ensure_ascii=False) + "\n"
