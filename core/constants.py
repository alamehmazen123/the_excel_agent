"""Engine-internal constants (no dependency on the app's root config)."""
from __future__ import annotations

# Display names of the sheets the engine appends to a workbook.
SHEET_INSIGHTS = "Insights"
SHEET_DASHBOARD = "Dashboard"
SHEET_PIVOT = "Pivot Analysis"
SHEET_KPI = "KPI Analysis"
SHEET_SUMMARY = "Executive Summary"
SHEET_SMART = "Smart Tables"


def output_sheet_names() -> list[str]:
    """All sheet names the engine may create -- skipped when re-analyzing."""
    return [SHEET_INSIGHTS, SHEET_DASHBOARD, SHEET_PIVOT, SHEET_KPI,
            SHEET_SUMMARY, SHEET_SMART]


def ordered_output_layout() -> list[str]:
    """Canonical left-to-right tab order for the analysis sheets, placed AFTER
    the data sheet(s):  KPI -> Pivot -> Smart Tables -> Insights -> Executive
    Summary -> Dashboard. Used by both the openpyxl writer and the Excel COM
    finalizer so the workbook always reads in this order."""
    return [SHEET_KPI, SHEET_PIVOT, SHEET_SMART, SHEET_INSIGHTS,
            SHEET_SUMMARY, SHEET_DASHBOARD]
