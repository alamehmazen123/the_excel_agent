"""Auto-computed hospital KPIs derived from raw data columns.

The engine detects when the right raw columns exist in a workbook and
auto-computes derived metrics like Occupancy Rate, Average Length of Stay,
Revenue per Patient, etc. These appear as additional entries on the KPI
scorecard, Insights sheet, and in Smart Tables alongside the raw measures.

Each derived metric knows its formula (for display), its dependencies (which
raw columns it needs), the report types it applies to, and how to compute it
from a row dict.

Design is pure ``core/`` — no UI, no Excel. Gracefully degrades: if a required
column is absent, the metric simply doesn't appear.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable, Optional

from .models import ColumnProfile, TableProfile, WorkbookProfile
from .semantic import MetricKind, ReportType, SemanticModel


@dataclass
class DerivedMetric:
    """A computed KPI that lives alongside raw measures on the output sheets."""
    name: str                          # Display name, e.g. "Occupancy Rate"
    unit: str                          # "percent", "number", "days", "currency"
    formula: str                       # Human-readable formula, e.g. "Bed-Days ÷ Available Bed-Days"
    description: str                   # What this tells a manager
    applies_to: list[ReportType]       # Which report types this is meaningful for
    # The raw column names this metric depends on. All must exist with valid data.
    depends_on: list[str] = field(default_factory=list)
    # Optional semantic dependency (a detected MetricKind column)
    depends_on_kind: Optional[MetricKind] = None

    def compute(self, row: dict[str, Any],
                col_map: dict[str, ColumnProfile]) -> Optional[float]:
        """Override in subclass. Returns the computed value or None."""
        raise NotImplementedError


class OccupancyRate(DerivedMetric):
    """Occupancy = Total Bed-Days Used / (Available Beds × Days in Period)."""
    def __init__(self) -> None:
        super().__init__(
            name="Occupancy Rate",
            unit="percent",
            formula="Bed-Days ÷ (Available Beds × Days)",
            description="Percentage of hospital bed capacity that was occupied",
            applies_to=[ReportType.CENSUS],
            depends_on=["bed_days", "available_beds"],
        )

    def compute(self, row: dict[str, Any],
                col_map: dict[str, ColumnProfile]) -> Optional[float]:
        used = row.get("bed_days")
        avail_beds = row.get("available_beds")
        days = row.get("days_in_period", 30)
        if used is None or avail_beds is None or avail_beds == 0:
            return None
        capacity = float(avail_beds) * float(days)
        return float(used) / capacity if capacity > 0 else None


class AverageLengthOfStay(DerivedMetric):
    """ALOS = Total Patient-Days / Number of Discharges."""
    def __init__(self) -> None:
        super().__init__(
            name="Average Length of Stay (ALOS)",
            unit="days",
            formula="Total Patient-Days ÷ Discharges",
            description="Average number of days a patient stays in hospital",
            applies_to=[ReportType.CENSUS],
            depends_on=["patient_days", "discharges"],
        )

    def compute(self, row: dict[str, Any],
                col_map: dict[str, ColumnProfile]) -> Optional[float]:
        days = row.get("patient_days")
        discharges = row.get("discharges")
        if days is None or discharges is None or discharges == 0:
            return None
        return float(days) / float(discharges)


class RevenuePerPatient(DerivedMetric):
    """Revenue per Patient = Total Revenue / Patient Count."""
    def __init__(self) -> None:
        super().__init__(
            name="Revenue per Patient",
            unit="currency",
            formula="Total Revenue ÷ Patient Count",
            description="Average revenue generated per patient encounter",
            applies_to=[ReportType.FINANCIAL],
            depends_on_kind=MetricKind.REVENUE,
            depends_on=["patient_count"],
        )

    def compute(self, row: dict[str, Any],
                col_map: dict[str, ColumnProfile]) -> Optional[float]:
        revenue_col = col_map.get("_revenue_col")
        revenue = row.get(revenue_col.name if revenue_col else "")
        patients = row.get("patient_count")
        if revenue is None or patients is None or patients == 0:
            return None
        return float(revenue) / float(patients)


class CostPerCase(DerivedMetric):
    """Cost per Case = Total Cost / Case Count."""
    def __init__(self) -> None:
        super().__init__(
            name="Cost per Case",
            unit="currency",
            formula="Total Cost ÷ Number of Cases",
            description="Average cost incurred per medical case",
            applies_to=[ReportType.FINANCIAL],
            depends_on_kind=MetricKind.COST,
            depends_on=["case_count"],
        )

    def compute(self, row: dict[str, Any],
                col_map: dict[str, ColumnProfile]) -> Optional[float]:
        cost_col = col_map.get("_cost_col")
        cost = row.get(cost_col.name if cost_col else "")
        cases = row.get("case_count")
        if cost is None or cases is None or cases == 0:
            return None
        return float(cost) / float(cases)


class BadDebtRatio(DerivedMetric):
    """Bad Debt Ratio = Bad Debt Write-offs / Gross Revenue."""
    def __init__(self) -> None:
        super().__init__(
            name="Bad Debt Ratio",
            unit="percent",
            formula="Bad Debt ÷ Gross Revenue",
            description="Portion of revenue written off as uncollectable",
            applies_to=[ReportType.FINANCIAL, ReportType.RECEIVABLES],
            depends_on=["bad_debt"],
            depends_on_kind=MetricKind.REVENUE,
        )

    def compute(self, row: dict[str, Any],
                col_map: dict[str, ColumnProfile]) -> Optional[float]:
        bd = row.get("bad_debt")
        rev_col = col_map.get("_revenue_col")
        rev = row.get(rev_col.name if rev_col else "")
        if bd is None or rev is None or rev == 0:
            return None
        return float(bd) / float(rev)


class CollectionRate(DerivedMetric):
    """Collection Rate = Amount Collected / Amount Billed."""
    def __init__(self) -> None:
        super().__init__(
            name="Collection Rate",
            unit="percent",
            formula="Collected ÷ Billed",
            description="Percentage of billed amounts successfully collected",
            applies_to=[ReportType.FINANCIAL, ReportType.RECEIVABLES],
            depends_on=["collections", "billed"],
        )

    def compute(self, row: dict[str, Any],
                col_map: dict[str, ColumnProfile]) -> Optional[float]:
        collected = row.get("collections")
        billed = row.get("billed")
        if collected is None or billed is None or billed == 0:
            return None
        return float(collected) / float(billed)


class BedTurnover(DerivedMetric):
    """Bed Turnover = Discharges / Available Beds."""
    def __init__(self) -> None:
        super().__init__(
            name="Bed Turnover",
            unit="number",
            formula="Discharges ÷ Available Beds",
            description="How many times each bed was used and freed in the period",
            applies_to=[ReportType.CENSUS],
            depends_on=["discharges", "available_beds"],
        )

    def compute(self, row: dict[str, Any],
                col_map: dict[str, ColumnProfile]) -> Optional[float]:
        discharges = row.get("discharges")
        beds = row.get("available_beds")
        if discharges is None or beds is None or beds == 0:
            return None
        return float(discharges) / float(beds)


# Registry of all available derived metrics (extend this list as new metrics are added).
ALL_DERIVED_METRICS: list[DerivedMetric] = [
    OccupancyRate(),
    AverageLengthOfStay(),
    RevenuePerPatient(),
    CostPerCase(),
    BadDebtRatio(),
    CollectionRate(),
    BedTurnover(),
]


def find_applicable_metrics(semantic: SemanticModel,
                            table: TableProfile) -> list[tuple[DerivedMetric, float]]:
    """Detect which derived metrics can be computed from this workbook's columns.

    Returns ``[(metric, total_value), ...]`` where ``total_value`` is the
    metric computed across ALL rows (for the scorecard). Returns empty list
    when nothing applies.
    """
    # Build a column name → profile map for quick lookup.
    col_map: dict[str, ColumnProfile] = {c.name: c for c in table.columns}

    # Detect semantic column assignments.
    rev = semantic.revenue
    cost = semantic.cost
    if rev is not None:
        col_map["_revenue_col"] = rev.column
    if cost is not None:
        col_map["_cost_col"] = cost.column

    # Heuristically map raw column names to the derived metric's ``depends_on``
    # slot names. We match by keyword against the header (both raw and library
    # meaning).
    header_keywords = {
        "bed_days": ("bed day", "bed-day", "beddays", "inpatient day", "patient day", "occupied bed"),
        "available_beds": ("available bed", "staffed bed", "licensed bed", "bed count", "bed capacity"),
        "patient_days": ("patient day", "inpatient day", "total day", "census day"),
        "discharges": ("discharge", "separation", "alos", "live discharge"),
        "patient_count": ("patient", "encounter", "admission", "visit", "case load", "census count"),
        "case_count": ("case", "procedure", "surgery", "operation", "encounter"),
        "bad_debt": ("bad debt", "write-off", "write off", "uncollectible", "uncollectable"),
        "collections": ("collection", "collected", "cash receipt", "payment received"),
        "billed": ("billed", "charge", "invoice", "gross revenue", "billing"),
        "days_in_period": ("days in period", "period days", "month days", "calendar days"),
    }

    # For each slot, find the best-matching column.
    slot_map: dict[str, Optional[ColumnProfile]] = {}
    for slot, keywords in header_keywords.items():
        best: Optional[ColumnProfile] = None
        best_score = 0
        for c in table.columns:
            if c.is_decoded_helper:
                continue
            name_lower = c.name.lower()
            score = sum(2 for kw in keywords if kw in name_lower)
            if score > best_score:
                best_score = score
                best = c
        slot_map[slot] = best

    # Now evaluate each derived metric.
    results: list[tuple[DerivedMetric, float]] = []
    for dm in ALL_DERIVED_METRICS:
        if semantic.report_type not in dm.applies_to:
            continue

        # Check all depends_on slots are resolved.
        all_resolved = True
        for dep in dm.depends_on:
            if dep not in slot_map or slot_map[dep] is None:
                all_resolved = False
                break
        if not all_resolved:
            continue

        # Check semantic dependency (e.g. needs a REVENUE column).
        if dm.depends_on_kind is not None:
            if dm.depends_on_kind == MetricKind.REVENUE and rev is None:
                continue
            if dm.depends_on_kind == MetricKind.COST and cost is None:
                continue

        # Compute the metric across ALL rows to get the total/aggregate value.
        total_value = 0.0
        count = 0
        for row in table.rows:
            val = dm.compute(row, col_map)
            if val is not None:
                total_value += val
                count += 1

        if count > 0:
            results.append((dm, total_value / count))  # average

    return results
