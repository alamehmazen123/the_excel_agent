"""The insight engine: turn a profiled workbook into ranked, typed findings.

Where the analyzers lay out numbers, this package decides which numbers MATTER
and says why — variance, concentration, anomalies, trend/forecast, ageing — as
explainable statistics that need no training data and run offline on a single
workbook.
"""
from __future__ import annotations

from .models import Insight, InsightKind, Severity
from .engine import detect_insights

__all__ = ["Insight", "InsightKind", "Severity", "detect_insights"]
