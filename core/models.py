"""Plain dataclasses describing what the engine discovers and produces.

These types are the stable contract between the engine and any front-end.
They contain no openpyxl/pandas objects so they are cheap to pass across
threads or serialize over an API boundary later.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Callable, Optional


class ColumnType(str, Enum):
    NUMERIC = "numeric"
    CURRENCY = "currency"
    PERCENT = "percent"
    DATE = "date"
    CATEGORICAL = "categorical"
    TEXT = "text"          # high-cardinality free text / IDs
    IDENTIFIER = "identifier"   # row id / ID / EVENT ID -- never a measure
    EMPTY = "empty"


# Header keywords used to rank "value" measures (money/amount) above prices etc.
_VALUE_KEYWORDS = ("pnl", "usdt", "usd", "amount", "profit", "revenue", "sales",
                   "income", "net", "gross", "value", "total", "balance", "$")


@dataclass
class ColumnProfile:
    name: str
    index: int                       # 0-based position within the table
    ctype: ColumnType
    count: int = 0                   # non-null values
    nulls: int = 0
    distinct: int = 0
    # The source cell number format (e.g. '"$"#,##0'), copied onto outputs.
    number_format: str = "General"
    # Numeric stats (None for non-numeric columns)
    minimum: Optional[float] = None
    maximum: Optional[float] = None
    mean: Optional[float] = None
    total: Optional[float] = None
    # For categorical/date columns: list of (value, count), most frequent first
    top_values: list[tuple[Any, int]] = field(default_factory=list)

    @property
    def is_measure(self) -> bool:
        # IDENTIFIER is numeric but is a row id -- never a measure.
        return self.ctype in (ColumnType.NUMERIC, ColumnType.CURRENCY, ColumnType.PERCENT)

    @property
    def is_value(self) -> bool:
        """A money/amount measure (not a percentage)."""
        return self.ctype in (ColumnType.NUMERIC, ColumnType.CURRENCY)

    @property
    def is_dimension(self) -> bool:
        return self.ctype in (ColumnType.CATEGORICAL, ColumnType.DATE)

    @property
    def value_score(self) -> int:
        """Higher = more likely the meaningful value to total (PNL over price)."""
        n = self.name.lower()
        score = sum(2 for k in _VALUE_KEYWORDS if k in n)
        if self.ctype == ColumnType.CURRENCY:
            score += 3
        if "price" in n:          # de-prioritise prices vs PNL/amounts
            score -= 2
        return score


@dataclass
class TableProfile:
    """A detected rectangular data region within a worksheet."""
    sheet_name: str
    header_row: int                  # 1-based row index of the header
    first_data_row: int              # 1-based
    last_data_row: int               # 1-based
    first_col: int                   # 1-based
    last_col: int                    # 1-based
    columns: list[ColumnProfile] = field(default_factory=list)
    # The actual data as a list of row dicts {column_name: value}, for analyzers.
    rows: list[dict[str, Any]] = field(default_factory=list)

    @property
    def row_count(self) -> int:
        return len(self.rows)

    @property
    def measures(self) -> list[ColumnProfile]:
        return [c for c in self.columns if c.is_measure]

    @property
    def value_measures(self) -> list[ColumnProfile]:
        """Money/amount measures, best candidate first (PNL before price)."""
        vals = [c for c in self.columns if c.is_value]
        return sorted(vals, key=lambda c: c.value_score, reverse=True)

    @property
    def percent_measures(self) -> list[ColumnProfile]:
        return [c for c in self.columns if c.ctype == ColumnType.PERCENT]

    @property
    def primary_value_measure(self) -> Optional[ColumnProfile]:
        vals = self.value_measures
        return vals[0] if vals else None

    @property
    def key_measures(self) -> list[ColumnProfile]:
        """Meaningful measures for tiles/stats: values first, then percents."""
        return self.value_measures + self.percent_measures

    def column(self, name: str) -> Optional[ColumnProfile]:
        for c in self.columns:
            if c.name == name:
                return c
        return None

    def measures_for(self, names: list[str]) -> list[ColumnProfile]:
        """Chosen measures (preserving the given order), or auto key_measures."""
        if names:
            chosen = [self.column(n) for n in names]
            chosen = [c for c in chosen if c is not None and c.is_measure]
            if chosen:
                return chosen
        return self.key_measures

    def value_for(self, name: Optional[str]) -> Optional[ColumnProfile]:
        if name:
            c = self.column(name)
            if c is not None and c.is_value:
                return c
        return self.primary_value_measure

    @property
    def dimensions(self) -> list[ColumnProfile]:
        # categorical dimensions (dates handled separately for grouping)
        return [c for c in self.columns if c.ctype == ColumnType.CATEGORICAL]

    @property
    def pivot_dimensions(self) -> list[ColumnProfile]:
        """Columns that make MEANINGFUL pivot row groupers.

        Includes clean categoricals AND moderate/high-card text (e.g.
        TRIGGER DETAIL) so the user can break down by them -- but excludes:
          * single-value columns (distinct < 2): a 1-row pivot is pointless;
          * near-unique columns (ratio > 0.95): row ids / GUIDs / free text.
        Wide dimensions are flagged via ``is_wide_dimension`` so the pivot can
        be limited to a Top-N by value.
        """
        out: list[ColumnProfile] = []
        for c in self.columns:
            if c.ctype not in (ColumnType.CATEGORICAL, ColumnType.TEXT):
                continue
            if c.distinct < 2:
                continue
            ratio = c.distinct / max(1, c.count)
            if ratio > 0.95:
                continue
            out.append(c)
        return out

    @staticmethod
    def is_wide_dimension(col: ColumnProfile) -> bool:
        return col.distinct > 25

    @property
    def identifier_column(self) -> Optional[ColumnProfile]:
        for c in self.columns:
            if c.ctype == ColumnType.IDENTIFIER:
                return c
        return None

    @property
    def date_columns(self) -> list[ColumnProfile]:
        return [c for c in self.columns if c.ctype == ColumnType.DATE]


@dataclass
class WorkbookProfile:
    path: str
    sheet_names: list[str] = field(default_factory=list)
    tables: list[TableProfile] = field(default_factory=list)
    primary_table_index: int = 0
    warnings: list[str] = field(default_factory=list)
    # Sheets that already contain a PivotTable -- detected and left untouched.
    pivot_sheets: list[str] = field(default_factory=list)
    # Set in Custom mode so the KPI/Dashboard/Summary sheets use the user's picks.
    preferred_value_name: Optional[str] = None
    preferred_measure_names: list[str] = field(default_factory=list)

    @property
    def primary(self) -> Optional[TableProfile]:
        if not self.tables:
            return None
        return self.tables[self.primary_table_index]


@dataclass
class MeasureChoice:
    """A value column the user chose to analyze, with a display format."""
    name: str
    # 'auto' | 'number' | 'usd' | 'lbp' | 'percent'
    format_kind: str = "auto"


@dataclass
class CustomSelection:
    """User-driven selection from the Custom Generate wizard."""
    sheet_name: Optional[str] = None
    dimensions: list[str] = field(default_factory=list)        # singles (one pivot each)
    measures: list[MeasureChoice] = field(default_factory=list)
    # Each inner list is a set of titles to nest into ONE combined pivot
    # (period x titles, Sum + % of total). The singles above are still produced.
    combinations: list[list[str]] = field(default_factory=list)

    def is_valid(self) -> bool:
        return bool(self.measures)


@dataclass
class AnalysisOptions:
    """Which analyses the user asked for (mirrors the UI checkboxes)."""
    dashboard: bool = True
    pivot: bool = True
    kpi: bool = True
    executive_summary: bool = True
    # Library-decoded plain summary tables (appears only when the library has
    # knowledge this workbook uses; see core/library + analyzers/smart_tables).
    smart_tables: bool = True
    # When set, the Custom Generate wizard drives the pivots/measures.
    custom: Optional["CustomSelection"] = None
    # If True, add a USD column (= LBP / 90000) next to each LBP value column.
    add_dollar: bool = False

    def any_selected(self) -> bool:
        return any((self.dashboard, self.pivot, self.kpi,
                    self.executive_summary, self.smart_tables))


@dataclass
class AnalysisResult:
    output_path: str
    sheets_created: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    summary_used_llm: bool = False
    notes: list[str] = field(default_factory=list)


# A progress callback: (fraction 0.0-1.0, human-readable status message).
ProgressCallback = Callable[[float, str], None]


def _noop_progress(fraction: float, message: str) -> None:  # pragma: no cover
    pass
