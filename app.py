"""Streamlit entry point. Run with: streamlit run app.py"""

from pathlib import Path
import sys

# Allow running from a checkout without installing the package first.
_src = Path(__file__).resolve().parent / "src"
if _src.is_dir() and str(_src) not in sys.path:
    sys.path.insert(0, str(_src))

from excel_report_maker.ui import main  # noqa: E402

main()
