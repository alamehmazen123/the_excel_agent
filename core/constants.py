"""Engine-internal constants (no dependency on the app's root config)."""
from __future__ import annotations

# Display names of the sheets the engine appends to a workbook.
SHEET_DASHBOARD = "Dashboard"
SHEET_PIVOT = "Pivot Analysis"
SHEET_KPI = "KPI Analysis"
SHEET_SUMMARY = "Executive Summary"


def output_sheet_names() -> list[str]:
    """All sheet names the engine may create -- skipped when re-analyzing."""
    return [SHEET_DASHBOARD, SHEET_PIVOT, SHEET_KPI, SHEET_SUMMARY]
