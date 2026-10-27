"""Config loading/validation tests."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from excel_report_maker.config import (
    list_templates,
    load_data,
    load_json_file,
    load_template,
    parse_json_text,
    strip_private_keys,
    validate_template,
)
from excel_report_maker.errors import DataError, TemplateError

ROOT = Path(__file__).resolve().parents[1]


def test_all_bundled_templates_load_and_validate() -> None:
    for name in ("staff_checklist", "compliance_checklist", "evaluation_matrix", "evaluation_scorecard"):
        template = load_template(name)
        assert template["report_type"] in ("checklist", "matrix")
        assert template["columns"]
    assert set(list_templates()) >= {
        "staff_checklist",
        "compliance_checklist",
        "evaluation_matrix",
        "evaluation_scorecard",
        "master_reference_config",
    }


def test_template_by_path_and_missing_template() -> None:
    path = ROOT / "configs" / "templates" / "staff_checklist.json"
    assert load_template(path)["sheet_name"] == "Admin & Compliance"
    with pytest.raises(FileNotFoundError):
        load_template("definitely_not_a_template")


def test_json5_comments_parse() -> None:
    parsed = parse_json_text(
        '{\n  // a comment\n  "report_type": "checklist",\n  /* multi */\n  "x": 1\n}', source="test"
    )
    assert parsed["report_type"] == "checklist"


def test_load_json_file_names_the_file_on_error(tmp_path: Path) -> None:
    bad = tmp_path / "bad.json"
    bad.write_text("{ not json", encoding="utf-8")
    with pytest.raises(ValueError, match="bad.json"):
        load_json_file(bad)


def test_strip_private_keys_only_removes_underscore_prefixed() -> None:
    assert strip_private_keys({"_comment": "x", "bold": True, "_x": 1}) == {"bold": True}


def test_load_data_samples_match_their_templates() -> None:
    checklist = load_template("staff_checklist")
    rows = load_data(ROOT / "data" / "samples" / "staff_checklist_data.json", checklist)
    assert len(rows) == 3
    matrix = load_template("evaluation_matrix")
    sections = load_data(ROOT / "data" / "samples" / "evaluation_matrix_data.json", matrix)
    assert len(sections) == 2
    assert sections[0]["header"].startswith("Component A")


def test_validate_template_rejects_empty_columns_and_bad_score_column() -> None:
    with pytest.raises(TemplateError, match="columns"):
        validate_template({"report_type": "checklist", "sheet_name": "S", "columns": []})
    with pytest.raises(TemplateError, match="score_column"):
        validate_template(
            {
                "report_type": "matrix",
                "sheet_name": "S",
                "columns": ["A", "B"],
                "score_column": 5,
            }
        )


def test_load_data_rejects_mismatched_payload(tmp_path: Path) -> None:
    template = load_template("staff_checklist")
    bad = tmp_path / "bad.json"
    bad.write_text(json.dumps([["only", "two"]]), encoding="utf-8")
    with pytest.raises(DataError):
        load_data(bad, template)
