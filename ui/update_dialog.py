"""Update-available dialog: shows version/date/notes, downloads, installs, restarts."""
from __future__ import annotations

import config
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QDialog, QHBoxLayout, QLabel, QMessageBox,
                               QProgressBar, QPushButton, QTextEdit, QVBoxLayout)

from core import updater

from .update_worker import UpdateDownloadWorker, run_on_thread


class UpdateDialog(QDialog):
    def __init__(self, info: updater.UpdateInfo, parent=None) -> None:
        super().__init__(parent)
        self._info = info
        self._thread = None
        self._worker = None
        self.setWindowTitle("Update Available")
        self.setMinimumWidth(460)
        self._build()

    def _build(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 22, 24, 20)
        layout.setSpacing(12)

        title = QLabel("A new version is available")
        title.setObjectName("TitleLabel")
        layout.addWidget(title)

        when = f" (released {self._info.release_date})" if self._info.release_date else ""
        ver = QLabel(f"Version {self._info.version}{when}\n"
                     f"You have version {config.APP_VERSION}.")
        ver.setObjectName("SubtitleLabel")
        layout.addWidget(ver)

        if self._info.notes:
            notes = QTextEdit()
            notes.setReadOnly(True)
            notes.setPlainText(self._info.notes)
            notes.setMaximumHeight(120)
            layout.addWidget(QLabel("What's new:"))
            layout.addWidget(notes)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        self.status = QLabel("")
        self.status.setObjectName("StatusLabel")
        layout.addWidget(self.status)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        self.later_btn = QPushButton("Later")
        self.later_btn.clicked.connect(self.reject)
        self.update_btn = QPushButton("Update Now")
        self.update_btn.setObjectName("PrimaryButton")
        self.update_btn.clicked.connect(self._start_download)
        buttons.addWidget(self.later_btn)
        buttons.addWidget(self.update_btn)
        layout.addLayout(buttons)

    # -- download / install ---------------------------------------------------
    def _start_download(self) -> None:
        if not updater.is_frozen():
            QMessageBox.information(
                self, config.APP_NAME,
                "Self-update only works in the installed app, not when running "
                "from source.")
            return
        self.update_btn.setEnabled(False)
        self.later_btn.setEnabled(False)
        self.progress.setVisible(True)
        self.status.setText("Downloading update…")

        self._worker = UpdateDownloadWorker(self._info)
        self._worker.progress.connect(self._on_progress)
        self._worker.ready.connect(self._on_ready)
        self._worker.failed.connect(self._on_failed)
        self._thread = run_on_thread(self._worker)

    def _on_progress(self, done: int, total: int) -> None:
        if total > 0:
            self.progress.setValue(int(done / total * 100))
            self.status.setText(f"Downloading… {done // 1_000_000} of {total // 1_000_000} MB")
        else:
            self.progress.setRange(0, 0)  # indeterminate when size unknown

    def _on_ready(self, path: str) -> None:
        if self._thread:
            self._thread.quit()
        self.status.setText("Installing and restarting…")
        QMessageBox.information(
            self, config.APP_NAME,
            "The update is ready. The app will now close and reopen on the new "
            "version.")
        try:
            updater.apply_update_and_restart(path)  # exits the process
        except Exception as exc:                     # noqa: BLE001
            self._on_failed(str(exc))

    def _on_failed(self, msg: str) -> None:
        if self._thread:
            self._thread.quit()
        self.progress.setVisible(False)
        self.update_btn.setEnabled(True)
        self.later_btn.setEnabled(True)
        self.status.setText("")
        QMessageBox.warning(
            self, config.APP_NAME,
            f"The update could not be installed:\n\n{msg}\n\n"
            "You can keep using the current version and try again later.")
