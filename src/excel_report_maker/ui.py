"""Streamlit UI for Excel Report Maker.

Three workflows share one engine: a guided Simple Builder, an
Advanced Builder that pairs a saved template with a data payload, and
a Template Architect for designing new templates. All sample people
in this UI are fictional.
"""

from __future__ import annotations

import io
import json
import re
from pathlib import Path
from typing import Any

import streamlit as st

from excel_report_maker.config import dump_json, load_json_file, parse_json_text, validate_template
from excel_report_maker.engine import generate_report
from excel_report_maker.errors import DataError, TemplateError

MODE_OVERVIEW = "🏠 Workspace Overview"
MODE_SIMPLE = "📋 Simple Builder (Guided)"
MODE_ADVANCED = "🚀 Advanced Builder"
MODE_ARCHITECT = "🛠️ Template Architect"
NAV_OPTIONS = [MODE_OVERVIEW, MODE_SIMPLE, MODE_ADVANCED, MODE_ARCHITECT]

TEMPLATE_DIRS = [Path("configs/templates"), Path("configs")]
DATA_DIRS = [Path("data/samples"), Path("local_data")]


def _set_mode(mode: str) -> None:
    st.session_state["nav_mode"] = mode


def _discover(directory_roots: list[Path]) -> dict[str, Path]:
    """JSON files under the given roots, keyed by display path.

    Keyed by relative path, not basename: in v1 two files with the
    same basename in different directories silently overwrote each
    other in the option list.
    """
    found: dict[str, Path] = {}
    for root in directory_roots:
        if root.is_dir():
            for path in sorted(root.glob("*.json")):
                found[str(path)] = path
    return found


def _get_report_bytes(template: dict[str, Any], data: Any) -> bytes:
    buffer = io.BytesIO()
    generate_report(template, data, buffer)
    return buffer.getvalue()


def _display_report_preview(template: dict[str, Any], data: Any) -> None:
    st.markdown("### 👀 Data Preview")
    report_type = template.get("report_type", "checklist")
    cols = template.get("columns", [])
    if report_type == "checklist":
        try:
            rows = [dict(zip(cols, row)) for row in data] if isinstance(data, list) else []
            st.table(rows)
        except Exception as exc:  # noqa: BLE001 - preview only, show the raw data
            st.warning(f"Could not generate tabular preview: {exc}")
            st.json(data)
        return
    # Matrix preview, one table per section. Malformed sections are
    # shown raw instead of crashing the page (v1 called .get on a list).
    sections = data if isinstance(data, list) else []
    for section in sections:
        if not isinstance(section, dict):
            st.write(section)
            continue
        st.markdown(f"**Section: {section.get('header', 'Unnamed')}**")
        try:
            rows = [dict(zip(cols, row)) for row in section.get("rows", [])]
            st.table(rows)
        except Exception:  # noqa: BLE001 - preview only
            st.write(section.get("rows", []))


def _download_section(state_key: str, name_key: str) -> None:
    """Render a download button for bytes stored by a Generate click.

    The button lives outside the Generate button's conditional: in v1
    it was nested inside it, so the rerun triggered by clicking
    Download removed the button before the download could happen.
    """
    payload = st.session_state.get(state_key)
    name = st.session_state.get(name_key)
    if payload and name:
        st.download_button(
            label=f"📥 Download {name}",
            data=payload,
            file_name=name,
            use_container_width=True,
            key=f"download-{state_key}",
        )


def _overview_page() -> None:
    st.subheader("🏠 Workspace Overview")
    st.markdown(
        "Welcome to the **Excel Report Generator**. Three workflows turn "
        "data into professionally formatted Excel documents. Select a "
        "method below or use the sidebar to begin."
    )
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("### 📋 Simple Builder")
        st.info("**Best for Quick Turnarounds**")
        st.markdown(
            "The Guided Wizard walks you through report creation step-by-step. "
            "No pre-defined JSON files required.\n\n"
            "- Define columns on the fly\n- Manual entry\n- Real-time data mapping"
        )
        st.button("Open Simple Builder", use_container_width=True, on_click=_set_mode, args=(MODE_SIMPLE,))
    with col2:
        st.markdown("### 🚀 Advanced Builder")
        st.success("**Best for Repetitive Reports**")
        st.markdown(
            "Separate structural templates from data. Ideal for bulk "
            "generation using existing configurations.\n\n"
            "- Load custom JSON templates\n- Use saved data payloads\n- Bulk processing support"
        )
        st.button("Open Advanced Builder", use_container_width=True, on_click=_set_mode, args=(MODE_ADVANCED,))
    with col3:
        st.markdown("### 🛠️ Template Architect")
        st.warning("**Best for Power Users**")
        st.markdown(
            "Design and save new structural templates with advanced Excel "
            "styling and validation rules.\n\n"
            "- Custom style definitions\n- Column-level validation\n- Save/export configurations"
        )
        st.button("Open Architect", use_container_width=True, on_click=_set_mode, args=(MODE_ARCHITECT,))

    st.divider()
    st.markdown("### 📖 Documentation Snippet")
    with st.expander("How does the Engine work?"):
        st.markdown(
            "1. **Template Engine:** Reads a JSON/JSON5 configuration defining columns, "
            "styles, and Excel settings (sheet names, widths, freeze panes).\n"
            "2. **Data Processor:** Validates the input data against the template's "
            "columns before anything is written.\n"
            "3. **XlsxWriter Integration:** Generates a streamable `.xlsx` file with "
            "frozen panes, conditional formatting, and data validation applied."
        )


def _advanced_page() -> None:
    st.subheader("🚀 Advanced Report Builder")
    st.markdown(
        "Use pre-defined **Structural Templates** and **Data Payloads** to generate "
        "complex reports. This mode is ideal for power users who want to reuse "
        "existing configurations or perform bulk generation."
    )
    st.divider()

    st.markdown("#### 🟦 1. Select Structural Template")
    template_options = _discover(TEMPLATE_DIRS)
    if not template_options:
        st.warning("No template files found under configs/templates/ or configs/.")
        return
    t_col1, t_col2 = st.columns([1, 1])
    with t_col1:
        selected_template = st.selectbox("Choose a template structure", list(template_options))
    template_config: dict[str, Any] = {}
    if selected_template:
        try:
            template_config = validate_template(load_json_file(template_options[selected_template]))
        except (ValueError, TemplateError) as exc:
            st.error(f"Template could not be loaded: {exc}")
            return
        with t_col2:
            st.info(
                f"**Loaded:** `{template_config.get('report_type', 'unknown')}` "
                f"with {len(template_config.get('columns', []))} columns."
            )
            with st.expander("🔍 Inspect Schema"):
                st.json(template_config)

    st.divider()
    st.markdown("#### 🟩 2. Provide Data Payload")
    d_col1, d_col2 = st.columns([1, 1])
    data_payload: Any = None
    with d_col1:
        data_mode = st.radio("Input Source", ["Sample File", "Manual (JSON5)"], horizontal=True)

    if data_mode == "Sample File":
        data_options = _discover(DATA_DIRS)
        with d_col1:
            if not data_options:
                st.warning("No sample data files found under data/samples/ or local_data/.")
            else:
                selected_data_file = st.selectbox("Select existing dataset", list(data_options))
                if selected_data_file:
                    try:
                        data_payload = load_json_file(data_options[selected_data_file])
                    except ValueError as exc:
                        st.error(f"Data file could not be loaded: {exc}")
        with d_col2:
            if data_payload is not None:
                st.success(f"**Data Ready:** {len(data_payload)} records/sections detected.")
                with st.expander("📝 Preview Raw Data"):
                    st.json(data_payload)
    else:
        placeholder = '[ ["Row1 Col1", "Row1 Col2"] ]'
        if template_config:
            cols = template_config.get("columns", [])
            if template_config.get("report_type", "checklist") == "checklist":
                placeholder = json.dumps([["Sample " + c for c in cols]], indent=4)
            else:
                placeholder = json.dumps(
                    [{"header": "Sample Section", "rows": [["Data " + str(i) for i in range(len(cols))]]}],
                    indent=4,
                )
        with d_col1:
            st.button(
                "Load Template Placeholder",
                on_click=lambda: st.session_state.update(manual_json_text=placeholder),
            )
            raw_data = st.text_area("Paste Data JSON List", height=200, key="manual_json_text")
        with d_col2:
            if raw_data.strip():
                try:
                    data_payload = parse_json_text(raw_data, source="pasted data")
                    st.success("**JSON Validated:** Content parsed successfully.")
                except ValueError as exc:
                    st.error(f"Invalid Data JSON: {exc}")

    st.divider()
    st.markdown("#### 🟨 3. Analysis & Preview")
    if template_config and data_payload is not None:
        m1, m2 = st.columns([1, 1])
        with m1:
            st.markdown("**Expected Columns**")
            st.code("\n".join(template_config.get("columns", [])))
        with m2:
            st.markdown("**Data Preview**")
            _display_report_preview(template_config, data_payload)

        st.divider()
        st.markdown("#### 📥 4. Generate & Export")
        if st.button("Generate Final Excel Report", use_container_width=True):
            try:
                st.session_state["advanced_bytes"] = _get_report_bytes(template_config, data_payload)
                st.session_state["advanced_name"] = template_config.get("file_name", "Generated_Report.xlsx")
                st.balloons()
            except (TemplateError, DataError, ValueError) as exc:
                st.error(f"Generation Error: {exc}")
        _download_section("advanced_bytes", "advanced_name")
    else:
        st.warning("Awaiting configuration and data input to enable generation.")


def _simple_page() -> None:
    st.subheader("📋 Simple Builder (Guided)")
    st.markdown(
        "### ⚡ Overview\n"
        "The **Simple Builder** is an interactive generator for rapid prototyping "
        "and human-first data entry. Define your layout with plain headers, fill "
        "in three sample rows, and download the formatted file.\n\n"
        "**How to use:** 1. Set the report title and export filename. "
        "2. List columns separated by commas. 3. Fill in the sample entries. "
        "4. Generate and download."
    )
    st.divider()

    defaults = {
        "simple_title": "Staff Compliance & Readiness Report",
        "simple_filename": "Compliance_Snapshot.xlsx",
        "active_headers": "Expert Name, Country, Background Check, Training Status, Expected Deployment",
        "sample_entry_1": ["Alex Johnson", "Canada", "Cleared", "Completed", "2026-03-01"],
        "sample_entry_2": ["Linda Chen", "Singapore", "Pending", "In Progress", "2026-04-15"],
        "sample_entry_3": ["Robert Smith", "United Kingdom", "Cleared", "Waitlisted", "2026-05-20"],
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, list(value) if isinstance(value, list) else value)

    def _reset_simple() -> None:
        for key, value in defaults.items():
            st.session_state[key] = list(value) if isinstance(value, list) else value
        for key in list(st.session_state):
            if key.startswith(("s1_field_", "s2_field_", "s3_field_")):
                del st.session_state[key]
        st.session_state.pop("simple_bytes", None)
        st.session_state.pop("simple_name", None)

    c1, c2 = st.columns([3, 1])
    with c1:
        st.markdown("### 🟦 Step 1: Branding & Identity")
        st.caption("Define how your document will be titled and saved.")
    with c2:
        st.button("🧹 Reset to Defaults", use_container_width=True, on_click=_reset_simple)

    is_valid = True
    col1, col2 = st.columns(2)
    with col1:
        title = st.text_input("Internal Report Title", key="simple_title")
        if not title.strip():
            st.warning("⚠️ **Required:** Please provide an internal report title.")
            is_valid = False
    with col2:
        file_name_input = st.text_input("Export Filename (.xlsx)", key="simple_filename")
        if file_name_input.strip():
            file_name = file_name_input.strip()
            if not file_name.lower().endswith(".xlsx"):
                file_name += ".xlsx"
        else:
            st.warning("⚠️ **Required:** Please provide a filename.")
            file_name = "report.xlsx"
            is_valid = False

    st.markdown("### 🟩 Step 2: Define Structure")
    st.caption("Comma-separated column headers. These are formatted and styled automatically.")
    st.text_area("Column Headers", height=68, key="active_headers")
    header_list = [h.strip() for h in st.session_state["active_headers"].split(",") if h.strip()]
    if not header_list:
        header_list = ["Column 1", "Column 2"]

    st.markdown("### 🟨 Step 3: Populate Content")
    st.caption("Edit the sample fields. The JSON mapping on the right updates in real time.")
    demo_cols = st.columns([1, 1.5])
    entries: list[list[str]] = []
    with demo_cols[0]:
        st.write("**Human-Readable Entry**")
        tabs = st.tabs(["🟢 Person A", "🟡 Person B", "🟠 Person C"])
        for tab_index, tab in enumerate(tabs, start=1):
            state_key = f"sample_entry_{tab_index}"
            with tab:
                st.markdown("---")
                entry: list[str] = []
                for i, header in enumerate(header_list):
                    current = st.session_state[state_key][i] if i < len(st.session_state[state_key]) else ""
                    st.markdown(f"#### 🏷️ {i}. {header}")
                    entry.append(
                        st.text_input(
                            f"field_{i}",
                            value=current,
                            key=f"s{tab_index}_field_{i}",
                            label_visibility="collapsed",
                        )
                    )
                st.session_state[state_key] = entry
                entries.append(entry)
    sync_data = [entry[: len(header_list)] for entry in entries]

    with demo_cols[1]:
        st.write("**Engine Data Mapping (JSON)**")

        def _format_mapped_json(data_rows: list[list[str]], headers: list[str]) -> str:
            mapped_rows = [{f"{idx} ({headers[idx]})": val for idx, val in enumerate(row)} for row in data_rows]
            parts = [json.dumps(row, indent=4).replace("\n", "\n    ") for row in mapped_rows]
            labels = ["🟢 Personnel Row 0", "🟡 Personnel Row 1", "🟠 Personnel Row 2"]
            body = ",\n\n    ".join(f"/* {label} */\n    {part}" for label, part in zip(labels, parts, strict=True))
            return f"[\n    {body}\n]"

        st.code(_format_mapped_json(sync_data, header_list), language="json")
        st.info(
            "📚 **Note:** Standard JSON does not support comments (the `/* ... */` lines). "
            "This illustration shows the indexed mapping only; the engine itself "
            "reads plain JSON/JSON5 data, and JSON5 files may carry comments."
        )
        st.caption("The engine combines the individual entries into a single structured array.")

    st.divider()
    st.markdown("### 🟦 Step 4: Validate & Export")
    st.caption("This grid shows how your data will look inside the final Excel file.")
    try:
        st.table([dict(zip(header_list, row)) for row in sync_data])
    except Exception:  # noqa: BLE001 - preview only
        st.json(sync_data)

    template = {
        "report_type": "checklist",
        "file_name": file_name,
        "sheet_name": "Data",
        "columns": header_list,
        "header_format": {"bold": True, "bg_color": "#2F5597", "font_color": "#FFFFFF"},
    }
    if st.button("Generate Simple Report", use_container_width=True, disabled=not is_valid):
        try:
            st.session_state["simple_bytes"] = _get_report_bytes(template, sync_data)
            st.session_state["simple_name"] = file_name
            st.balloons()
        except (TemplateError, DataError, ValueError) as exc:
            st.error(f"Generation Error: {exc}")
    _download_section("simple_bytes", "simple_name")


def _safe_template_name(raw: str) -> str:
    """A bare template filename: no directories, no traversal, .json suffix."""
    name = Path(raw.strip()).name
    name = re.sub(r"[^A-Za-z0-9_.\-]", "_", name)
    if name.startswith("."):
        name = name.lstrip(".")
    if not name:
        raise ValueError("Template name must not be empty.")
    if not name.lower().endswith(".json"):
        name += ".json"
    if name in {".json", "..json"}:
        raise ValueError("Template name must not be empty.")
    return name


def _architect_page() -> None:
    st.subheader("🛠️ Template Architect")
    st.info("Design a new structural template. Use the 'Master Reference' as a guide for valid stylings.")

    with st.expander("📚 Common Style Snippets (from Master Spec)"):
        st.code(
            '"header_format": {\n'
            '    "bold": true, "bg_color": "#2F5597", "font_color": "#FFFFFF", "border": 1, "align": "center"\n'
            "},\n"
            '"column_formats": [\n'
            '    { "index": 0, "width": 20, "bold": true, "suffix": ": " },\n'
            '    { "index": 1, "validation": { "validate": "list", "source": ["A", "B", "C"] } }\n'
            "]",
            language="json",
        )

    col1, col2 = st.columns(2)
    with col1:
        t_name = st.text_input("Template Name", "new_report_template")
        t_type = st.selectbox("Report Type", ["checklist", "matrix"])
        t_sheet = st.text_input("Sheet Name", "Sheet1")
    with col2:
        t_cols = st.text_area("Columns (one per line)", "ID\nName\nStatus")
        t_file = st.text_input("Default Output Filename", f"{t_name}.xlsx")

    columns_list = [c.strip() for c in t_cols.split("\n") if c.strip()]
    boilerplate: dict[str, Any] = {
        "report_type": t_type,
        "file_name": t_file,
        "sheet_name": t_sheet,
        "columns": columns_list,
        "header_format": {"bold": True, "bg_color": "#4472C4", "font_color": "#FFFFFF"},
        "column_formats": [{"index": i, "width": 15} for i, _ in enumerate(columns_list)],
        "conditional_formats": [],
    }
    if t_type == "matrix":
        boilerplate["score_column"] = len(columns_list) - 2 if len(columns_list) >= 2 else 0
        boilerplate["width_cap"] = 100

    st.markdown("### 🎛️ Refine Structural JSON")
    refined_json = st.text_area(
        "Customize Layout & Logic", value=dump_json(boilerplate), height=300, key="architect_json"
    )
    try:
        current_config = validate_template(parse_json_text(refined_json, source="template draft"))
        if current_config.get("report_type") == "checklist":
            dummy_data: Any = [["Sample " + str(i) for i in range(len(current_config.get("columns", [])))]]
        else:
            dummy_data = [
                {
                    "header": "Sample Section",
                    "rows": [["Sample " + str(i) for i in range(len(current_config.get("columns", [])))]],
                }
            ]
        _display_report_preview(current_config, dummy_data)
    except (ValueError, TemplateError, DataError) as exc:
        st.warning(f"Preview unavailable until the draft is valid: {exc}")

    if st.button("💾 Save Template"):
        try:
            final_config = validate_template(parse_json_text(refined_json, source="template draft"))
            safe_name = _safe_template_name(t_name)
            save_dir = Path("configs/templates")
            save_dir.mkdir(parents=True, exist_ok=True)
            save_path = save_dir / safe_name
            with save_path.open("w", encoding="utf-8") as handle:
                handle.write(dump_json(final_config))
            st.success(f"Template saved to {save_path}!")
        except (ValueError, TemplateError, OSError) as exc:
            st.error(f"Failed to save template: {exc}")


def main() -> None:
    st.set_page_config(page_title="Excel Report Generator", layout="wide")
    st.markdown(
        """
        <style>
        [data-testid="stHeaderActionElements"] { display: none !important; }
        [data-testid="stSidebarContent"] {
            padding-top: 2rem !important;
            padding-left: 1rem !important;
            padding-right: 1rem !important;
        }
        .block-container {
            padding-left: 3rem !important;
            padding-right: 3rem !important;
            padding-top: 1.5rem !important;
        }
        textarea { resize: none !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.sidebar.title("📊 Excel Report Generator")
    mode = st.sidebar.radio("Select Input Method", NAV_OPTIONS, key="nav_mode")
    if mode == MODE_OVERVIEW:
        _overview_page()
    elif mode == MODE_ADVANCED:
        _advanced_page()
    elif mode == MODE_SIMPLE:
        _simple_page()
    elif mode == MODE_ARCHITECT:
        _architect_page()
    st.divider()
    st.info(
        "The **Advanced Builder** strictly separates styling (Templates) from content (Data) "
        "for professional workflows."
    )


if __name__ == "__main__":
    main()
