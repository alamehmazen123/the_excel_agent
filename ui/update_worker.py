"""Threaded wrappers around core.updater so the UI never blocks."""
from __future__ import annotations

import config
from PySide6.QtCore import QObject, QThread, Signal

from core import updater


class UpdateCheckWorker(QObject):
    """Checks the manifest in the background; emits found(UpdateInfo) if newer."""
    found = Signal(object)        # updater.UpdateInfo
    none = Signal()

    def run(self) -> None:
        info = updater.check_for_update(config.UPDATE_MANIFEST_URL, config.APP_VERSION)
        if info is not None:
            self.found.emit(info)
        else:
            self.none.emit()


class UpdateDownloadWorker(QObject):
    """Downloads the new exe, reporting progress; emits ready(path) or failed(msg)."""
    progress = Signal(int, int)   # bytes_done, bytes_total
    ready = Signal(str)           # path to downloaded exe
    failed = Signal(str)

    def __init__(self, info) -> None:
        super().__init__()
        self._info = info

    def run(self) -> None:
        try:
            path = updater.download_update(self._info, self.progress.emit)
            self.ready.emit(path)
        except Exception as exc:                  # noqa: BLE001
            self.failed.emit(str(exc))


def run_on_thread(worker: QObject, start_method_name: str = "run") -> QThread:
    """Move ``worker`` to a fresh QThread and start it. Caller keeps both refs."""
    thread = QThread()
    worker.moveToThread(thread)
    thread.started.connect(getattr(worker, start_method_name))
    thread.start()
    return thread
