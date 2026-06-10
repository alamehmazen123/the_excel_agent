"""Declarative plan of the PivotTables to build, consumed by the COM finalizer.

The plan is a SCENARIO GENERATOR derived from column roles, so it adapts to any
workbook and produces only *meaningful* pivots:

  * per date column (grouped Month+Year): trade count + total value;
  * per groupable dimension (SYMBOL / ACTION / EXIT REASON / STRATEGY SOURCE /
    TRIGGER DETAIL ...): total value in $ and total percent in %;
  * date x dimension breakdowns (e.g. month/year x STRATEGY SOURCE);
  * a Measure Statistics pivot on the KPI sheet.

Row-id columns are never used; single-value and near-unique columns are skipped;
wide dimensions are limited to a Top-N by value so they stay readable.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from .aggregate import group_sum
from .constants import SHEET_KPI, SHEET_PIVOT
from .models import ColumnProfile, TableProfile, WorkbookProfile

# Excel PivotField summary-function constants (xlConsolidationFunction).
XL_SUM = -4157
XL_AVERAGE = -4106
XL_COUNT = -4112
XL_MIN = -4139
XL_MAX = -4136

INT_FORMAT = "#,##0"
# PNL PCT etc. are already in percent units (0.11 == 0.11%), so append a literal
# % rather than Excel's "%" format (which would multiply by 100).
PCT_FORMAT = '0.00"%"'
DEFAULT_CURRENCY = '"$"#,##0.00'

TOP_N = 20                 # cap for wide dimensions (e.g. TRIGGER DETAIL)
MAX_DIMS = 8               # don't generate an unbounded number of pivots
MAX_CROSS_DIMS = 5         # date x dimension combos


@dataclass
class DataFieldSpec:
    source_field: str
    func: int
    caption: str
    number_format: str


@dataclass
class PivotSpec:
    target_sheet: str
    title: str
    row_fields: list[str] = field(default_factory=list)
    data_fields: list[DataFieldSpec] = field(default_factory=list)
    group_date_field: Optional[str] = None
    sort_field: Optional[str] = None     # AutoSort the last row field by this caption
    # When set, only these labels of the last row field are shown (Top-N by value).
    visible_items: Optional[list[str]] = None


def _currency_format(col: Optional[ColumnProfile]) -> str:
    if col is None:
        return DEFAULT_CURRENCY
    f = col.number_format or "General"
    if any(sym in f for sym in ("$", "€", "£", "¥", "₹", "[$", "USD")):
        return f
    return DEFAULT_CURRENCY


def _percent_format(col: ColumnProfile) -> str:
    """If the source already uses an Excel % format the value is a fraction
    (0.11 == 11%), so keep '%'. Otherwise the number is already a percent
    (0.11 == 0.11%), so append a literal % without multiplying."""
    if "%" in (col.number_format or ""):
        return "0.00%"
    return PCT_FORMAT


def _top_labels(table: TableProfile, dim: ColumnProfile,
                measure: ColumnProfile, n: int) -> list[str]:
    """Top-N labels of ``dim`` ranked by |sum of measure| -- computed offline."""
    ranked = group_sum(table, dim, measure, top_n=10_000)
    ranked.sort(key=lambda kv: abs(kv[1]), reverse=True)
    return [str(k) for k, _ in ranked[:n]]


def build_pivot_plan(profile: WorkbookProfile) -> list[PivotSpec]:
    table = profile.primary
    if table is None:
        return []

    value = table.primary_value_measure
    pct = table.percent_measures[0] if table.percent_measures else None
    dims = table.pivot_dimensions[:MAX_DIMS]
    dates = table.date_columns[:3]
    id_col = table.identifier_column
    count_field = id_col.name if id_col else (
        table.columns[0].name if table.columns else None)
    vfmt = _currency_format(value)

    if value is None and pct is None:
        return []

    plan: list[PivotSpec] = []
    value_caption = f"Total {value.name}" if value else None

    # --- per-date: trade count + total value, grouped Month + Year -----------
    for d in dates:
        dfs: list[DataFieldSpec] = []
        if count_field:
            dfs.append(DataFieldSpec(count_field, XL_COUNT, "Record Count", INT_FORMAT))
        if value is not None:
            dfs.append(DataFieldSpec(value.name, XL_SUM, value_caption, vfmt))
        if dfs:
            plan.append(PivotSpec(SHEET_PIVOT, f"By {d.name} (Month/Year)",
                                  [d.name], dfs, group_date_field=d.name))

    # --- per groupable dimension: total value ($) and total percent (%) ------
    for dim in dims:
        wide = TableProfile.is_wide_dimension(dim)
        if value is not None:
            plan.append(PivotSpec(
                SHEET_PIVOT, f"{value.name} by {dim.name}", [dim.name],
                [DataFieldSpec(value.name, XL_SUM, value_caption, vfmt)],
                sort_field=value_caption,
                visible_items=_top_labels(table, dim, value, TOP_N) if wide else None))
        if pct is not None:
            pct_caption = f"Total {pct.name}"
            plan.append(PivotSpec(
                SHEET_PIVOT, f"{pct.name} by {dim.name}", [dim.name],
                [DataFieldSpec(pct.name, XL_SUM, pct_caption, _percent_format(pct))],
                sort_field=pct_caption,
                visible_items=_top_labels(table, dim, pct, TOP_N) if wide else None))

    # --- date x dimension breakdowns (e.g. month/year x STRATEGY SOURCE) -----
    if dates and value is not None:
        primary_date = dates[0].name
        for dim in dims[:MAX_CROSS_DIMS]:
            wide = TableProfile.is_wide_dimension(dim)
            plan.append(PivotSpec(
                SHEET_PIVOT,
                f"{value.name} by {dates[0].name} (Month/Year) & {dim.name}",
                [primary_date, dim.name],
                [DataFieldSpec(value.name, XL_SUM, value_caption, vfmt)],
                group_date_field=primary_date,
                sort_field=value_caption,
                visible_items=_top_labels(table, dim, value, TOP_N) if wide else None))

    # --- KPI sheet: Measure Statistics on the primary value measure ----------
    if value is not None:
        plan.append(PivotSpec(
            SHEET_KPI, "Measure Statistics", [],
            [
                DataFieldSpec(value.name, XL_SUM, value_caption, vfmt),
                DataFieldSpec(value.name, XL_AVERAGE, f"Average {value.name}", vfmt),
                DataFieldSpec(value.name, XL_COUNT, "Count", INT_FORMAT),
                DataFieldSpec(value.name, XL_MIN, f"Min {value.name}", vfmt),
                DataFieldSpec(value.name, XL_MAX, f"Max {value.name}", vfmt),
            ]))

    return plan
