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
    EMPTY = "empty"


@dataclass
class ColumnProfile:
    name: str
    index: int                       # 0-based position within the table
    ctype: ColumnType
    count: int = 0                   # non-null values
    nulls: int = 0
    distinct: int = 0
    # Numeric stats (None for non-numeric columns)
    minimum: Optional[float] = None
    maximum: Optional[float] = None
    mean: Optional[float] = None
    total: Optional[float] = None
    # For categorical/date columns: list of (value, count), most frequent first
    top_values: list[tuple[Any, int]] = field(default_factory=list)

    @property
    def is_measure(self) -> bool:
        return self.ctype in (ColumnType.NUMERIC, ColumnType.CURRENCY, ColumnType.PERCENT)

    @property
    def is_dimension(self) -> bool:
        return self.ctype in (ColumnType.CATEGORICAL, ColumnType.DATE)


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
    def dimensions(self) -> list[ColumnProfile]:
        return [c for c in self.columns if c.is_dimension]

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

    @property
    def primary(self) -> Optional[TableProfile]:
        if not self.tables:
            return None
        return self.tables[self.primary_table_index]


@dataclass
class AnalysisOptions:
    """Which analyses the user asked for (mirrors the UI checkboxes)."""
    dashboard: bool = True
    pivot: bool = True
    kpi: bool = True
    executive_summary: bool = True

    def any_selected(self) -> bool:
        return any((self.dashboard, self.pivot, self.kpi, self.executive_summary))


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
