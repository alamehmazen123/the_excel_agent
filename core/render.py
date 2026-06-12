"""Neutral, framework-free description of what a result sheet should contain.

Analyzers build :class:`SheetSpec` objects; the writer renders them with
openpyxl. This indirection keeps analyzers testable without Excel and lets a
future front-end render the same specs to HTML/PDF instead.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Optional


class ChartKind(str, Enum):
    BAR = "bar"
    LINE = "line"
    PIE = "pie"
    PARETO = "pareto"     # bars (value) + cumulative % line on a secondary axis
    COMBO = "combo"       # bars (actual) + line (trend/forecast) overlay


class NumberFormat(str, Enum):
    GENERAL = "general"
    INTEGER = "integer"
    DECIMAL = "decimal"
    CURRENCY = "currency"     # US dollars ("$#,##0.00")
    LBP = "lbp"               # Lebanese Pounds ("#,##0 LBP") — the hospital default
    PERCENT = "percent"
    DATE = "date"


@dataclass
class KpiTile:
    label: str
    value: str          # pre-formatted display string
    caption: str = ""   # optional sub-text (e.g. "+12.3% vs prior period")
    good: Optional[bool] = None  # True=green, False=red, None=neutral


@dataclass
class DataTable:
    title: str
    headers: list[str]
    rows: list[list[Any]]
    # Per-column number formats (len == len(headers)); GENERAL if omitted.
    formats: list[NumberFormat] = field(default_factory=list)
    # Smart-table visuals (0-based column indices): in-cell gradient data bars,
    # and a green→red color scale (heatmap). Empty = plain table.
    bar_columns: list[int] = field(default_factory=list)
    scale_columns: list[int] = field(default_factory=list)


@dataclass
class ChartSpec:
    kind: ChartKind
    title: str
    categories: list[Any]            # x-axis / slice labels
    series_name: str
    values: list[float]
    # Optional overlay line for PARETO (cumulative %) / COMBO (trend/forecast).
    line_values: Optional[list[float]] = None
    line_name: str = ""


@dataclass
class TextBlock:
    title: str
    paragraphs: list[str]
    # 'normal' | 'highlight' (bold navy) | 'recommend' (red bold, underlined title)
    style: str = "normal"


@dataclass
class SheetSpec:
    """Everything needed to render one output worksheet, top to bottom."""
    name: str
    heading: str
    subheading: str = ""
    kpi_tiles: list[KpiTile] = field(default_factory=list)
    tables: list[DataTable] = field(default_factory=list)
    charts: list[ChartSpec] = field(default_factory=list)
    text_blocks: list[TextBlock] = field(default_factory=list)
