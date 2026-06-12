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
from .models import (ColumnProfile, ColumnType, CustomSelection, TableProfile,
                     WorkbookProfile)

# Excel PivotField summary-function constants (xlConsolidationFunction).
XL_SUM = -4157
XL_AVERAGE = -4106
XL_COUNT = -4112
XL_MIN = -4139
XL_MAX = -4136
# Show-values-as calculation: % of grand total (XlPivotFieldCalculation).
XL_PERCENT_OF_TOTAL = 8

INT_FORMAT = "#,##0"
PERCENT_OF_TOTAL_FORMAT = "0.00%"
# PNL PCT etc. are already in percent units (0.11 == 0.11%), so append a literal
# % rather than Excel's "%" format (which would multiply by 100).
PCT_FORMAT = '0.00"%"'
# Hospital reports are in Lebanese Pounds by default. Use $ only when the header
# or the source cell format explicitly indicates dollars.
LBP_FORMAT = '#,##0" LBP"'
USD_FORMAT = '"$"#,##0.00'
USD_PER_LBP = 90000          # exchange rate: 90,000 LBP = 1 USD

TOP_N = 20                 # cap for wide dimensions (e.g. TRIGGER DETAIL)
MAX_DIMS = 8               # don't generate an unbounded number of pivots
MAX_CROSS_DIMS = 5         # date x dimension combos


@dataclass
class DataFieldSpec:
    source_field: str
    func: int
    caption: str
    number_format: str
    # Optional "show values as" calculation, e.g. XL_PERCENT_OF_TOTAL.
    calculation: Optional[int] = None
    # Optional pivot CalculatedField formula (e.g. "='Revenue'/90000" for USD).
    # When set, source_field is the NEW calculated field's name.
    calc_formula: Optional[str] = None


def _usd_pair(col_name: str) -> DataFieldSpec:
    """A USD data field (= column / 90000) implemented as a pivot calculated field."""
    return DataFieldSpec(f"{col_name} (USD)", XL_SUM, f"Total {col_name} ($)",
                         USD_FORMAT, calc_formula=f"='{col_name}'/{USD_PER_LBP}")


def value_fields(col: ColumnProfile, fmt: str, cap: str,
                 add_dollar: bool) -> list[DataFieldSpec]:
    """The native value field, plus a USD-converted field when requested and the
    column isn't already in dollars."""
    fields = [DataFieldSpec(col.name, XL_SUM, cap, fmt)]
    if add_dollar and not _is_dollar(col):
        fields.append(_usd_pair(col.name))
    return fields


@dataclass
class PivotSpec:
    target_sheet: str
    title: str
    row_fields: list[str] = field(default_factory=list)
    data_fields: list[DataFieldSpec] = field(default_factory=list)
    group_date_field: Optional[str] = None
    # Full label order (highest value first) for the last categorical row field;
    # applied deterministically via PivotItem.Position (Excel AutoSort is flaky).
    ordered_labels: Optional[list[str]] = None
    # When set, only these labels of the last row field are shown (Top-N by value).
    visible_items: Optional[list[str]] = None


def _is_dollar(col: Optional[ColumnProfile]) -> bool:
    """True only if the column is explicitly in dollars (header or cell format)."""
    if col is None:
        return False
    f = (col.number_format or "")
    if "$" in f or "USD" in f:
        return True
    n = col.name.lower()
    return ("$" in n or "usd" in n or "dollar" in n or "usdt" in n)


def _currency_format(col: Optional[ColumnProfile]) -> str:
    """Dollar format if the column is explicitly $, otherwise Lebanese Pounds."""
    if _is_dollar(col):
        f = col.number_format or ""
        return f if ("$" in f) else USD_FORMAT
    return LBP_FORMAT


def _percent_format(col: ColumnProfile) -> str:
    """If the source already uses an Excel % format the value is a fraction
    (0.11 == 11%), so keep '%'. Otherwise the number is already a percent
    (0.11 == 0.11%), so append a literal % without multiplying."""
    if "%" in (col.number_format or ""):
        return "0.00%"
    return PCT_FORMAT


# Format chosen by the user in the Custom wizard -> (Excel format, role).
def _resolve_choice(col: ColumnProfile, kind: str) -> tuple[str, str]:
    if kind == "usd":
        return '"$"#,##0.00', "value"
    if kind == "lbp":
        return '#,##0" LBP"', "value"
    if kind == "number":
        return "#,##0.00", "value"
    if kind == "percent":
        return _percent_format(col), "percent"
    # auto -> follow the detected type
    if col.ctype == ColumnType.PERCENT:
        return _percent_format(col), "percent"
    if col.ctype == ColumnType.CURRENCY:
        return _currency_format(col), "value"
    return "#,##0.00", "value"


def _ranked_labels(table: TableProfile, dim: ColumnProfile,
                   measure: ColumnProfile) -> list[str]:
    """All labels of ``dim`` ranked by sum of ``measure``, highest value first."""
    ranked = group_sum(table, dim, measure, top_n=10_000)
    ranked.sort(key=lambda kv: kv[1], reverse=True)
    return [str(k) for k, _ in ranked]


def _combined_spec(table: TableProfile, date_name: Optional[str],
                   dim_cols: list[ColumnProfile],
                   value_specs: list[tuple], pct_col: ColumnProfile,
                   pct_caption: str, add_dollar: bool) -> Optional[PivotSpec]:
    """One pivot nesting a date period + 2-3 titles, with Sum value(s) + % of total.

    ``value_specs`` is a list of (ColumnProfile, format, caption) tuples.
    """
    rows = ([date_name] if date_name else []) + [d.name for d in dim_cols]
    if len(rows) < 2:                 # 'combined' needs at least two titles
        return None
    data: list[DataFieldSpec] = []
    for (col, fmt, cap) in value_specs:
        data.extend(value_fields(col, fmt, cap, add_dollar))
    data.append(DataFieldSpec(pct_col.name, XL_SUM, f"{pct_col.name} (% of total)",
                              PERCENT_OF_TOTAL_FORMAT, calculation=XL_PERCENT_OF_TOTAL))
    last_dim = dim_cols[-1] if dim_cols else None
    ordered = _ranked_labels(table, last_dim, pct_col) if last_dim else None
    bits = ([f"{date_name} (Month/Year)"] if date_name else []) + [d.name for d in dim_cols]
    return PivotSpec(SHEET_PIVOT, f"{pct_caption} by " + " & ".join(bits) + " (combined)",
                     rows, data, group_date_field=date_name, ordered_labels=ordered)


def build_custom_plan(table: TableProfile, sel: CustomSelection,
                      add_dollar: bool = False) -> list[PivotSpec]:
    """Build pivots from the user's explicit dimension/measure picks."""
    # Resolve each chosen measure to (column, format, role, caption).
    resolved = []
    for mc in sel.measures:
        col = table.column(mc.name)
        if col is None:
            continue
        fmt, role = _resolve_choice(col, mc.format_kind)
        resolved.append((col, fmt, role, f"Total {col.name}"))
    if not resolved:
        return []
    value_measures = [(c, f, cap) for (c, f, role, cap) in resolved if role == "value"]

    dims = [table.column(n) for n in sel.dimensions]
    dims = [d for d in dims if d is not None]
    date_dims = [d for d in dims if d.ctype == ColumnType.DATE]
    cat_dims = [d for d in dims if d.ctype != ColumnType.DATE]

    id_col = table.identifier_column
    count_field = id_col.name if id_col else (
        table.columns[0].name if table.columns else None)

    plan: list[PivotSpec] = []

    # Date summaries: count + every chosen measure, grouped Month/Year.
    for d in date_dims:
        dfs = []
        if count_field:
            dfs.append(DataFieldSpec(count_field, XL_COUNT, "Record Count", INT_FORMAT))
        for (col, fmt, role, cap) in resolved:
            if role == "value":
                dfs.extend(value_fields(col, fmt, cap, add_dollar))
            else:
                dfs.append(DataFieldSpec(col.name, XL_SUM, cap, fmt))
        plan.append(PivotSpec(SHEET_PIVOT, f"By {d.name} (Month/Year)",
                              [d.name], dfs, group_date_field=d.name))

    # Each categorical dimension x each chosen measure.
    for d in cat_dims:
        wide = TableProfile.is_wide_dimension(d)
        for (col, fmt, role, cap) in resolved:
            ordered = _ranked_labels(table, d, col)
            data = (value_fields(col, fmt, cap, add_dollar) if role == "value"
                    else [DataFieldSpec(col.name, XL_SUM, cap, fmt)])
            plan.append(PivotSpec(
                SHEET_PIVOT, f"{col.name} by {d.name}", [d.name],
                data,
                ordered_labels=ordered,
                visible_items=ordered[:TOP_N] if wide else None))

    # Date x dimension cross-breakdowns (first date x each category, first value).
    if date_dims and value_measures:
        pd_name = date_dims[0].name
        vcol, vfmt, vcap = value_measures[0]
        for d in cat_dims[:MAX_CROSS_DIMS]:
            wide = TableProfile.is_wide_dimension(d)
            ordered = _ranked_labels(table, d, vcol)
            plan.append(PivotSpec(
                SHEET_PIVOT,
                f"{vcol.name} by {date_dims[0].name} (Month/Year) & {d.name}",
                [pd_name, d.name],
                value_fields(vcol, vfmt, vcap, add_dollar),
                group_date_field=pd_name,
                ordered_labels=ordered,
                visible_items=ordered[:TOP_N] if wide else None))

    # Combination pivots: each requested set of titles becomes ONE nested pivot
    # (period x titles, Sum of each value + % of total). Singles above remain.
    if value_measures:
        vcol, vfmt, vcap = value_measures[0]
        for combo in sel.combinations:
            cols = [table.column(n) for n in combo]
            cols = [c for c in cols if c is not None]
            cdate = next((c.name for c in cols if c.ctype == ColumnType.DATE), None)
            ccats = [c for c in cols if c.ctype != ColumnType.DATE]
            combined = _combined_spec(table, cdate, ccats[:3], value_measures,
                                      vcol, vcap, add_dollar)
            if combined is not None:
                plan.append(combined)

    # KPI Measure Statistics on the first chosen value measure.
    if value_measures:
        vcol, vfmt, vcap = value_measures[0]
        plan.append(PivotSpec(
            SHEET_KPI, "Measure Statistics", [],
            [DataFieldSpec(vcol.name, XL_SUM, vcap, vfmt),
             DataFieldSpec(vcol.name, XL_AVERAGE, f"Average {vcol.name}", vfmt),
             DataFieldSpec(vcol.name, XL_COUNT, "Count", INT_FORMAT),
             DataFieldSpec(vcol.name, XL_MIN, f"Min {vcol.name}", vfmt),
             DataFieldSpec(vcol.name, XL_MAX, f"Max {vcol.name}", vfmt)]))
    return plan


def build_pivot_plan(profile: WorkbookProfile,
                     custom: Optional[CustomSelection] = None,
                     add_dollar: bool = False) -> list[PivotSpec]:
    table = profile.primary
    if table is None:
        return []

    if custom is not None and custom.is_valid():
        return build_custom_plan(table, custom, add_dollar)

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
            dfs.extend(value_fields(value, vfmt, value_caption, add_dollar))
        if dfs:
            plan.append(PivotSpec(SHEET_PIVOT, f"By {d.name} (Month/Year)",
                                  [d.name], dfs, group_date_field=d.name))

    # --- per groupable dimension: total value ($) and total percent (%) ------
    for dim in dims:
        wide = TableProfile.is_wide_dimension(dim)
        if value is not None:
            ordered = _ranked_labels(table, dim, value)
            plan.append(PivotSpec(
                SHEET_PIVOT, f"{value.name} by {dim.name}", [dim.name],
                value_fields(value, vfmt, value_caption, add_dollar),
                ordered_labels=ordered,
                visible_items=ordered[:TOP_N] if wide else None))
        if pct is not None:
            pct_caption = f"Total {pct.name}"
            ordered = _ranked_labels(table, dim, pct)
            plan.append(PivotSpec(
                SHEET_PIVOT, f"{pct.name} by {dim.name}", [dim.name],
                [DataFieldSpec(pct.name, XL_SUM, pct_caption, _percent_format(pct))],
                ordered_labels=ordered,
                visible_items=ordered[:TOP_N] if wide else None))

    # --- date x dimension breakdowns (e.g. month/year x STRATEGY SOURCE) -----
    if dates and value is not None:
        primary_date = dates[0].name
        for dim in dims[:MAX_CROSS_DIMS]:
            wide = TableProfile.is_wide_dimension(dim)
            ordered = _ranked_labels(table, dim, value)
            plan.append(PivotSpec(
                SHEET_PIVOT,
                f"{value.name} by {dates[0].name} (Month/Year) & {dim.name}",
                [primary_date, dim.name],
                [DataFieldSpec(value.name, XL_SUM, value_caption, vfmt)],
                group_date_field=primary_date,
                ordered_labels=ordered,
                visible_items=ordered[:TOP_N] if wide else None))

    # --- combined: period x (pairs of titles), Sum value + % of total -------
    # Keep the singles above AND add combination pivots for the top dimension
    # pairs, so the user gets both and can delete whichever they don't need.
    if value is not None:
        import itertools  # noqa: PLC0415
        narrow = [d for d in dims if not TableProfile.is_wide_dimension(d)][:3]
        date_name = dates[0].name if dates else None
        for a, b in list(itertools.combinations(narrow, 2))[:3]:
            combined = _combined_spec(table, date_name, [a, b],
                                      [(value, vfmt, value_caption)], value,
                                      value_caption, add_dollar)
            if combined is not None:
                plan.append(combined)

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
