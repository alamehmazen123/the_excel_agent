"""Typed findings produced by the insight engine.

An :class:`Insight` is a single, self-contained statement about the data: a
headline a hospital manager can read, the precise numbers behind it, how serious
it is, and a pointer to the evidence (which dimension/measure/period proves it).
Everything downstream — the Insights sheet, the narrative, chart selection — is
driven by these objects rather than by re-deriving statistics ad hoc.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Optional


class InsightKind(str, Enum):
    VARIANCE = "variance"            # period-over-period change in a measure
    CONCENTRATION = "concentration"  # a few items dominate (Pareto)
    ANOMALY = "anomaly"              # an out-of-band point in a time series
    TREND = "trend"                  # sustained direction + simple forecast
    AGING = "aging"                  # receivables/balance ageing buckets
    LOSS = "loss"                    # negative / loss-making records
    LEADER = "leader"                # the top contributor in a dimension


class Severity(str, Enum):
    INFO = "info"
    WATCH = "watch"
    HIGH = "high"

    @property
    def rank(self) -> int:
        return {"info": 0, "watch": 1, "high": 2}[self.value]


@dataclass
class Insight:
    kind: InsightKind
    severity: Severity
    title: str                       # short headline, e.g. "Revenue fell 23%"
    detail: str                      # one-sentence explanation with the numbers
    # Machine-readable score in [0,1] used to RANK insights (higher = surface first).
    score: float = 0.0
    # Whether this is good (True), bad (False) or neutral (None) for the hospital.
    good: Optional[bool] = None
    # Evidence pointers so the sheet/narrative can reference or chart the source.
    measure: Optional[str] = None
    dimension: Optional[str] = None
    period: Optional[str] = None
    # Optional small payload for chart building (e.g. waterfall steps, buckets).
    evidence: dict[str, Any] = field(default_factory=dict)

    def sort_key(self) -> tuple[int, float]:
        return (self.severity.rank, self.score)
