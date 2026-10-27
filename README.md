# Excel Report Maker 📊

One template, one data file, one formatted workbook. Define how a report looks once, in JSON, then generate it as many times as you like with different data — a staff compliance checklist this morning, an evaluation matrix this afternoon, same engine, no copy-paste formatting.

Built for proposal and compliance workflows where the styling has to be identical every time: frozen panes, dropdown validation, conditional coloring, and sectioned score tables, all declared in the template instead of reapplied by hand.

## Quick start

```bash
pip install -e ".[ui]"

# Check a template and its data before generating anything
excel-report-maker validate --template staff_checklist --data data/samples/staff_checklist_data.json

# Generate a workbook
excel-report-maker generate --template staff_checklist \
  --data data/samples/staff_checklist_data.json \
  --output outputs/Staff_Compliance_Checklist.xlsx

# Or use the guided web UI
streamlit run app.py
```

The CLI also works from an installed package with no checkout at all: bundled templates are found by name (`staff_checklist`, `compliance_checklist`, `evaluation_matrix`, `evaluation_scorecard`, `master_reference_config`). Run `excel-report-maker list-templates` to see them.

## How it works

Styling and content are strictly separated:

- **Template** — columns, sheet name, header and cell formats, column widths and suffixes, dropdown validation, conditional formatting, freeze panes, and page setup. JSON, with JSON5 comments allowed.
- **Data** — rows for a checklist, or headed sections of rows for a matrix. JSON or JSON5.
- **Engine** — validates the data against the template before writing anything, then renders with XlsxWriter. A row with the wrong number of values is a clear error naming the row, not a crash deep in a spreadsheet library.

```text
excel-report-maker/
├── app.py                        # Streamlit entry point
├── src/excel_report_maker/
│   ├── engine.py                 # Checklist + matrix generators
│   ├── config.py                 # Template/data loading and validation
│   ├── cli.py                    # excel-report-maker command
│   ├── ui.py                     # Streamlit UI (Simple / Advanced / Architect)
│   └── templates/                # Bundled templates (also in configs/)
├── configs/templates/            # Editable copies of the bundled templates
├── configs/master_reference_config.json  # Feature showcase with embedded data
└── data/samples/                 # Fictional sample payloads
```

## Templates

| Template | Type | What it produces |
|---|---|---|
| `staff_checklist` | Checklist | 12-column compliance checklist with TECH-6/7/8 signing columns, dropdowns, and green/yellow status coloring |
| `compliance_checklist` | Checklist | 10-column staff compliance checklist with medical/security/contract status |
| `evaluation_matrix` | Matrix | Sectioned applicant scoring by position, with the total-score column highlighted |
| `evaluation_scorecard` | Matrix | Sectioned candidate scorecard with education/experience/interview weightings |
| `master_reference_config` | Checklist | Showcase of every engine feature: named style registry, explicit widths, page setup, validation, conditional formats |

Templates may name formats directly, or reference a `styles` registry and the built-ins `green`, `yellow`, `red`, and `alert_red`. Keys starting with `_` inside format and rule objects are documentation and are stripped before anything reaches XlsxWriter. Matrices declare their highlighted column with `score_column` and cap auto-fit widths with `width_cap`.

## The three UI workflows

- **Simple Builder (Guided)** — type headers and three sample rows, download the file. No JSON required.
- **Advanced Builder** — pick a saved template and a data file (or paste JSON5), preview the result, download it. This is the workflow for repetitive reports.
- **Template Architect** — design a new template in the browser and save it under `configs/templates/`.

Private working files belong in `local_configs/` and `local_data/`, and generated files in `outputs/`. All three are gitignored: real candidate and proposal data must never be committed. The samples shipped here are entirely fictional — the people, emails, and scores are invented for demonstration.

## Development

```bash
pip install -e ".[dev]"
ruff check src tests && ruff format --check src tests
mypy src
pytest -q
```

Tests generate real workbooks in temporary directories and read them back with openpyxl, checking headers, values, freeze panes, column widths, page setup, and conditional formatting — including the master reference config, whose conditional formats silently produced nothing in v1.

## Notes

- TECH-6/7/8 are standard procurement form labels; this tool is not affiliated with, or endorsed by, the World Bank or any other institution. The `compliance_checklist` and `evaluation_scorecard` templates are generic shapes, not official forms.
- Formerly `mp-excelMaker`. See [CHANGELOG.md](CHANGELOG.md).

## License

MIT — see [LICENSE](LICENSE).
