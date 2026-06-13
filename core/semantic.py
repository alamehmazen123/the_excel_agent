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

from collections import defaultdict
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Optional

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
    # Goal/purpose inferred from the account-code categories (the library knows
    # 713… = revenues, 601… = purchases, …). ``purpose`` is a human label shown
    # to the reader; ``purpose_kind`` is the metric nature it implies.
    purpose: str = ""
    purpose_kind: Optional[MetricKind] = None
    account_column: Optional[str] = None     # the code column the purpose came from
    category_totals: dict[str, float] = field(default_factory=dict)

    @property
    def is_revenue_report(self) -> bool:
        return self.purpose_kind == MetricKind.REVENUE

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


# Map an account-category banner to the metric nature it implies.
_EXPENSE_WORDS = ("expense", "purchase", "salary", "salaries", "tax",
                  "maintenance", "insurance", "fuel", "office", "marketing",
                  "donation", "bank charge", "general", "supplies")


def _nature_of_category(category: str) -> Optional[MetricKind]:
    c = (category or "").lower()
    if "revenue" in c or "income" in c or "sales" in c:
        return MetricKind.REVENUE
    if any(w in c for w in _EXPENSE_WORDS):
        return MetricKind.COST
    if "cash" in c or "bank account" in c:
        return MetricKind.BALANCE
    return None


_PURPOSE_LABEL = {
    MetricKind.REVENUE: "Revenue",
    MetricKind.COST: "Expenses",
    MetricKind.BALANCE: "Cash & balances",
}


def _distinct(table: TableProfile, col: ColumnProfile, limit: int = 300) -> list[Any]:
    out, seen = [], set()
    for row in table.rows:
        v = row.get(col.name)
        if v in (None, ""):
            continue
        k = str(v)
        if k in seen:
            continue
        seen.add(k)
        out.append(v)
        if len(out) >= limit:
            break
    return out


def _detect_purpose(table: TableProfile, library,
                    money: Optional[ColumnProfile]) -> Optional[tuple]:
    """Infer what the workbook is ABOUT from its account-code categories.

    Finds the code column whose library map carries categories, then sums the
    money (or row counts) per category. Returns (code_col_name, dominant_kind,
    purpose_label, category_totals) or None."""
    if library is None or getattr(library, "is_empty", True):
        return None
    best = None                       # (col, code_map, coverage)
    for col in table.columns:
        if col.is_decoded_helper or col.ctype not in (
                ColumnType.CATEGORICAL, ColumnType.TEXT, ColumnType.IDENTIFIER):
            continue
        vals = _distinct(table, col)
        cm = library.best_map_for_values(vals)
        if cm is None or not getattr(cm, "categories", None):
            continue
        cov = cm.coverage(vals)
        if best is None or cov > best[2]:
            best = (col, cm, cov)
    if best is None:
        return None

    col, cm, _cov = best
    totals: dict[str, float] = defaultdict(float)
    for row in table.rows:
        cat = cm.category_of(row.get(col.name))
        if not cat:
            continue
        amt = 1.0
        if money is not None:
            v = row.get(money.name)
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                amt = abs(float(v)) or 1.0
        totals[cat] += amt
    if not totals:
        return None

    # Group the per-category money into metric natures, pick the dominant one.
    by_nature: dict[MetricKind, float] = defaultdict(float)
    for cat, amt in totals.items():
        nat = _nature_of_category(cat)
        if nat is not None:
            by_nature[nat] += amt
    if not by_nature:
        return None
    dominant = max(by_nature, key=by_nature.get)
    label = _PURPOSE_LABEL.get(dominant, dominant.value.title())
    return col.name, dominant, label, dict(totals)


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

    model = SemanticModel(report_type=ReportType.GENERIC, measures=measures,
                          table=table)

    # Goal detection from account categories (overrides keyword guesses). The
    # money weighting uses the strongest money measure available.
    money = measures[0].column if measures and measures[0].kind.is_money else None
    if money is None:
        money = table.primary_value_measure
    purpose = _detect_purpose(table, library, money)
    if purpose is not None:
        acol, kind, label, totals = purpose
        model.purpose = label
        model.purpose_kind = kind
        model.account_column = acol
        model.category_totals = totals
        # Re-tag the money measures to the detected nature so revenue/expense
        # logic (RAG direction, sign handling, narrative) is correct.
        for m in measures:
            if m.kind.is_money:
                m.kind = kind
        model.report_type = ReportType.FINANCIAL
    else:
        kinds = [m.kind for m in measures]
        model.report_type = _detect_report_type(table, meanings, kinds)
    return model
