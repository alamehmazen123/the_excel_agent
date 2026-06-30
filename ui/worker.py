"""Background worker that runs the engine without freezing the UI."""
from __future__ import annotations

from PySide6.QtCore import QObject, QThread, Signal

from core.llm.groq_client import GroqNarrator
from core.models import AnalysisOptions, AnalysisResult
from core.pipeline import Engine

from . import settings_store


class AnalysisWorker(QObject):
    progress = Signal(float, str)        # fraction, message
    finished = Signal(object)            # AnalysisResult
    failed = Signal(str)                 # error message

    def __init__(self, workbook_path: str, options: AnalysisOptions) -> None:
        super().__init__()
        self._path = workbook_path
        self._options = options

    def run(self) -> None:
        try:
            narrator = None
            if self._options.executive_summary:
                key = settings_store.effective_key()
                if key:
                    narrator = GroqNarrator(key, model=settings_store.get_model())
            engine = Engine(narrator=narrator)
            result: AnalysisResult = engine.run(
                self._path, self._options, self._emit_progress)
            self.finished.emit(result)
        except Exception as exc:                       # noqa: BLE001
            self.failed.emit(str(exc))

    def _emit_progress(self, fraction: float, message: str) -> None:
        self.progress.emit(fraction, message)


def start_analysis(path: str, options: AnalysisOptions,
                   on_progress, on_finished, on_failed) -> tuple[QThread, AnalysisWorker]:
    """Wire a worker onto its own thread and start it. Caller keeps the refs."""
    thread = QThread()
    worker = AnalysisWorker(path, options)
    worker.moveToThread(thread)
    thread.started.connect(worker.run)
    worker.progress.connect(on_progress)
    worker.finished.connect(on_finished)
    worker.failed.connect(on_failed)
    worker.finished.connect(thread.quit)
    worker.failed.connect(thread.quit)
    thread.finished.connect(thread.deleteLater)
    thread.start()
    return thread, worker
