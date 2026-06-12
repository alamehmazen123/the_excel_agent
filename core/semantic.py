"""Semantic layer: teach the engine what a hospital workbook MEANS.

The profiler (``core/profiler.py``) classifies columns by *shape* (numeric,
currency, date, categorical …). This module adds a layer of *meaning* on top —
it decides whether a money column is REVENUE or COST, whether a numeric column
is a VOLUME (admissions, bed-days) or a RATIO (occupancy %), and what KIND of
report the whole workbook is (financial / receivables / census / generic).

That semantic read is what lets the insight engine and the Insights sheet talk
in hospital language ("net revenue fell", "receivables are ageing") instead of
generic "column 4 went down". It is pure ``core/`` logic — no UI, no Excel — and
it degrades gracefully: anything it can't classify becomes GENERIC and is still
analysed, just without the domain phrasing.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional

from .formatting import is_dollar_column
from .models import ColumnProfile, ColumnType, TableProfile, WorkbookProfile


class MetricKind(str, Enum):
    REVENUE = "revenue"     # money coming in (revenue/sales/charges/collections)
    COST = "cost"           # money going out (cost/expense/payable/salary)
    BALANCE = "balance"     # a standing amount (receivable/AR/outstanding/debt)
    VOLUME = "volume"       # a count (admissions/visits/cases/bed-days/units)
    RATIO = "ratio"         # a percentage/rate (occupancy %, margin, denial rate)
    AMOUNT = "amount"       # money with no clearer role
    GENERIC = "generic"     # numeric, role unknown

    @property
    def is_money(self) -> bool:
        return self in (MetricKind.REVENUE, MetricKind.COST,
                        MetricKind.BALANCE, MetricKind.AMOUNT)

    @property
    def higher_is_better(self) -> Optional[bool]:
        """Whether a rise in this metric is good (None = depends/unknown)."""
        if self == MetricKind.REVENUE:
            return True
        if self in (MetricKind.COST, MetricKind.BALANCE):
            return False
        return None


class ReportType(str, Enum):
    FINANCIAL = "financial"        # GL / P&L / chart-of-accounts
    RECEIVABLES = "receivables"    # AR / ageing / outstanding balances
    CENSUS = "census"              # admissions / discharges / LOS / occupancy
    OPERATIONS = "operations"      # OR / pharmacy / throughput volumes
    GENERIC = "generic"


# Keyword banks. Matched against BOTH the raw header and its library meaning, so
# they work whether the column is "REV" with a glossary entry or spelled out.
_REVENUE_KW = ("revenue", "sales", "income", "charge", "billing", "billed",
               "collection", "collected", "receipt", "gross", "net revenue",
               "turnover", "payment received", "earnings", "proceeds")
_COST_KW = ("cost", "expense", "expenditure", "salary", "wage", "payroll",
            "purchase", "supply", "supplies", "cogs", "spend", "payable",
            "disbursement", "overhead")
_BALANCE_KW = ("receivable", " ar ", "a/r", "outstanding", "due", "debt",
               "balance", "ageing", "aging", "arrears", "unpaid", "dso")
_VOLUME_KW = ("count", "qty", "quantity", "units", "number of", "visits",
              "visit", "admission", "discharge", "case", "cases", "patient",
              "encounter", "bed day", "bed-day", "beddays", "los", "length of stay",
              "occupanc", "procedure", "tests", "scans", "doses", "claims")
_RATIO_KW = ("rate", "ratio", "margin", "percent", "percentage", "pct",
             " % ", "occupancy", "utilization", "utilisation", "share")

_RECEIVABLES_KW = ("receivable", "ageing", "aging", "outstanding", "dso",
                   "days sales", "collection", "unpaid", "arrears")
_CENSUS_KW = ("admission", "discharge", "length of stay", "los", "occupanc",
              "bed", "census", "inpatient", "patient day")
_OPERATIONS_KW = ("operating room", " or ", "theatre", "pharmacy", "drug",
                  "throughput", "procedure", "scan", "imaging", "lab")
_FINANCIAL_KW = ("account", "ledger", "journal", "debit", "credit", "pnl",
                 "p&l", "profit", "expense", "revenue", "balance sheet")


def _haystack(col: ColumnProfile, meaning: str) -> str:
    return f" {col.name.lower()} | {meaning.lower()} "


def classify_metric(col: ColumnProfile, meaning: str = "") -> MetricKind:
    """Assign a hospital metric role to a numeric/percent column."""
    text = _haystack(col, meaning)

    if col.ctype == ColumnType.PERCENT:
        return MetricKind.RATIO
    if any(k in text for k in _RATIO_KW) and col.ctype != ColumnType.CURRENCY:
        return MetricKind.RATIO

    money = col.ctype == ColumnType.CURRENCY or is_dollar_column(col)
    if any(k in text for k in _BALANCE_KW):
        return MetricKind.BALANCE if money else MetricKind.BALANCE
    if any(k in text for k in _REVENUE_KW):
        return MetricKind.REVENUE
    if any(k in text for k in _COST_KW):
        return MetricKind.COST
    if any(k in text for k in _VOLUME_KW):
        return MetricKind.VOLUME

    if money:
        return MetricKind.AMOUNT
    if col.ctype == ColumnType.NUMERIC:
        # Integer-ish numeric with modest magnitude reads as a count.
        return MetricKind.VOLUME if _looks_like_count(col) else MetricKind.GENERIC
    return MetricKind.GENERIC


def _looks_like_count(col: ColumnProfile) -> bool:
    mx = col.maximum if col.maximum is not None else 0
    mn = col.minimum if col.minimum is not None else 0
    return mn >= 0 and mx <= 1_000_000 and (col.mean or 0) < 100_000


@dataclass
class MeasureSemantic:
    column: ColumnProfile
    kind: MetricKind
    meaning: str                 # library meaning, else the header

    @property
    def name(self) -> str:
        return self.column.name


@dataclass
class SemanticModel:
    """The engine's domain read of one workbook (its primary table)."""
    report_type: ReportType = ReportType.GENERIC
    measures: list[MeasureSemantic] = field(default_factory=list)
    table: Optional[TableProfile] = None

    # -- convenience accessors used by the insight engine ------------------- #
    def of_kind(self, *kinds: MetricKind) -> list[MeasureSemantic]:
        return [m for m in self.measures if m.kind in kinds]

    def kind_of(self, name: str) -> MetricKind:
        for m in self.measures:
            if m.name == name:
                return m.kind
        return MetricKind.GENERIC

    @property
    def primary_money(self) -> Optional[MeasureSemantic]:
        """The headline money measure: prefer revenue, then any amount/cost."""
        for pref in (MetricKind.REVENUE, MetricKind.AMOUNT,
                     MetricKind.COST, MetricKind.BALANCE):
            got = self.of_kind(pref)
            if got:
                # Within a kind, the profiler already ranks by value_score.
                return got[0]
        return self.measures[0] if self.measures else None

    @property
    def revenue(self) -> Optional[MeasureSemantic]:
        got = self.of_kind(MetricKind.REVENUE)
        return got[0] if got else None

    @property
    def cost(self) -> Optional[MeasureSemantic]:
        got = self.of_kind(MetricKind.COST)
        return got[0] if got else None

    @property
    def balance(self) -> Optional[MeasureSemantic]:
        got = self.of_kind(MetricKind.BALANCE)
        return got[0] if got else None


def _detect_report_type(table: TableProfile, meanings: dict[str, str],
                        kinds: list[MetricKind]) -> ReportType:
    # Build one big haystack of every header + its library meaning.
    blob = " | ".join(
        f"{c.name.lower()} {meanings.get(c.name, '').lower()}" for c in table.columns)

    def has(words: tuple[str, ...]) -> int:
        return sum(1 for w in words if w in blob)

    receivables = has(_RECEIVABLES_KW) + (2 if MetricKind.BALANCE in kinds else 0)
    census = has(_CENSUS_KW)
    operations = has(_OPERATIONS_KW)
    financial = has(_FINANCIAL_KW) + (1 if MetricKind.REVENUE in kinds
                                      and MetricKind.COST in kinds else 0)

    scores = {
        ReportType.RECEIVABLES: receivables,
        ReportType.CENSUS: census,
        ReportType.OPERATIONS: operations,
        ReportType.FINANCIAL: financial,
    }
    best = max(scores, key=lambda k: scores[k])
    return best if scores[best] >= 2 else ReportType.GENERIC


def analyze(profile: WorkbookProfile, library=None) -> SemanticModel:
    """Build a :class:`SemanticModel` for a workbook's primary table."""
    table = profile.primary
    if table is None:
        return SemanticModel()

    # Resolve each header's real meaning via the library (no-op if empty).
    meanings: dict[str, str] = {}
    if library is not None and not getattr(library, "is_empty", True):
        for c in table.columns:
            meanings[c.name] = library.meaning_of(c.name)

    measures: list[MeasureSemantic] = []
    # Rank money measures first (value_measures), then percents — same priority
    # the rest of the engine uses, so primary_money lines up with the pivots.
    for col in table.value_measures + table.percent_measures:
        if col.is_decoded_helper:
            continue
        meaning = meanings.get(col.name, col.name)
        measures.append(MeasureSemantic(col, classify_metric(col, meaning), meaning))

    kinds = [m.kind for m in measures]
    rtype = _detect_report_type(table, meanings, kinds)
    return SemanticModel(report_type=rtype, measures=measures, table=table)
