"""Analyzer abstract base class."""
from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Optional

from ..models import WorkbookProfile
from ..render import SheetSpec


class Analyzer(ABC):
    """One analysis that turns a profiled workbook into one output sheet."""

    #: Short key matching AnalysisOptions fields ("dashboard", "kpi", ...).
    key: str = ""
    #: Display name of the sheet this analyzer creates.
    sheet_name: str = ""

    @abstractmethod
    def applies_to(self, profile: WorkbookProfile) -> bool:
        """Return True if this analysis is meaningful for the given data.

        The UI uses this to disable options that can't produce useful output.
        """

    @abstractmethod
    def run(self, profile: WorkbookProfile) -> Optional[SheetSpec]:
        """Produce the sheet spec, or None if nothing could be computed."""
