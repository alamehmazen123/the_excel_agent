"""Main application window."""
from __future__ import annotations

import os
import subprocess

import config
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (QCheckBox, QFileDialog, QFrame, QGridLayout,
                               QHBoxLayout, QLabel, QMessageBox, QProgressBar,
                               QPushButton, QVBoxLayout, QWidget)

from core.models import AnalysisOptions
from core.pipeline import Engine

from .settings_dialog import SettingsDialog
from .update_dialog import UpdateDialog
from .update_worker import UpdateCheckWorker, run_on_thread
from .worker import start_analysis


class MainWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(config.APP_NAME)
        self.setMinimumSize(720, 640)
        self.setAcceptDrops(True)

        self._path: str = ""
        self._thread = None
        self._worker = None
        self._busy = False

        self._update_thread = None
        self._update_check = None

        self._build()
        self._refresh_buttons()
        self._check_for_updates()

    # -- layout ---------------------------------------------------------------
    def _build(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 24)
        root.setSpacing(18)

        # Header
        header = QHBoxLayout()
        titles = QVBoxLayout()
        title = QLabel(config.APP_NAME)
        title.setObjectName("TitleLabel")
        subtitle = QLabel("Turn any workbook into dashboards, pivots, KPIs and an executive summary.")
        subtitle.setObjectName("SubtitleLabel")
        titles.addWidget(title)
        titles.addWidget(subtitle)
        header.addLayout(titles)
        header.addStretch(1)
        gear = QPushButton("⚙  Settings")
        gear.clicked.connect(self._open_settings)
        header.addWidget(gear, alignment=Qt.AlignTop)
        root.addLayout(header)

        # Workbook selection card
        sel_card = self._card()
        sel = QVBoxLayout(sel_card)
        sel.setContentsMargins(20, 18, 20, 18)
        sel.addWidget(self._section("Workbook Selection"))
        row = QHBoxLayout()
        self.browse_btn = QPushButton("Browse Workbook…")
        self.browse_btn.clicked.connect(self._browse)
        row.addWidget(self.browse_btn)
        self.file_label = QLabel("No workbook selected.  (You can also drag a file here.)")
        self.file_label.setObjectName("FileLabel")
        row.addWidget(self.file_label, 1)
        self.clear_btn = QPushButton("✕")
        self.clear_btn.setObjectName("ClearButton")
        self.clear_btn.setToolTip("Remove the selected workbook")
        self.clear_btn.setFixedSize(34, 34)
        self.clear_btn.clicked.connect(self._clear_workbook)
        self.clear_btn.setVisible(False)
        row.addWidget(self.clear_btn)
        sel.addLayout(row)
        root.addWidget(sel_card)

        # Output mode card
        opt_card = self._card()
        opt = QVBoxLayout(opt_card)
        opt.setContentsMargins(20, 18, 20, 18)
        opt.addWidget(self._section("Output Mode"))
        grid = QGridLayout()
        grid.setHorizontalSpacing(40)
        self.cb_dashboard = QCheckBox("Dashboard")
        self.cb_pivot = QCheckBox("Pivot Analysis")
        self.cb_kpi = QCheckBox("KPI Analysis")
        self.cb_summary = QCheckBox("Executive Summary")
        for cb in (self.cb_dashboard, self.cb_pivot, self.cb_kpi, self.cb_summary):
            cb.setChecked(True)
        grid.addWidget(self.cb_dashboard, 0, 0)
        grid.addWidget(self.cb_pivot, 0, 1)
        grid.addWidget(self.cb_kpi, 1, 0)
        grid.addWidget(self.cb_summary, 1, 1)
        opt.addLayout(grid)
        root.addWidget(opt_card)

        # Actions
        actions = QHBoxLayout()
        self.analyze_btn = QPushButton("Analyze Workbook")
        self.analyze_btn.setObjectName("PrimaryButton")
        self.analyze_btn.clicked.connect(self._analyze)
        actions.addWidget(self.analyze_btn, 2)
        self.open_wb_btn = QPushButton("Open Workbook")
        self.open_wb_btn.clicked.connect(self._open_workbook)
        actions.addWidget(self.open_wb_btn, 1)
        self.open_folder_btn = QPushButton("Open Output Folder")
        self.open_folder_btn.clicked.connect(self._open_folder)
        actions.addWidget(self.open_folder_btn, 1)
        root.addLayout(actions)

        # Progress
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        root.addWidget(self.progress)
        self.status_label = QLabel("Ready.")
        self.status_label.setObjectName("StatusLabel")
        root.addWidget(self.status_label)
        root.addStretch(1)

        # Footer: version + release date (left) and update status (right).
        footer = QHBoxLayout()
        self.version_label = QLabel(self._version_text())
        self.version_label.setObjectName("FooterLabel")
        footer.addWidget(self.version_label)
        footer.addStretch(1)
        self.update_label = QLabel("")
        self.update_label.setObjectName("FooterLabel")
        footer.addWidget(self.update_label)
        root.addLayout(footer)

    def _version_text(self) -> str:
        return f"Version {config.APP_VERSION}  •  Released {config.BUILD_DATE}"

    def _card(self) -> QFrame:
        f = QFrame()
        f.setObjectName("Card")
        return f

    def _section(self, text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setObjectName("SectionLabel")
        return lbl

    # -- drag & drop ----------------------------------------------------------
    def dragEnterEvent(self, e) -> None:
        if e.mimeData().hasUrls() and self._is_excel(e.mimeData().urls()[0].toLocalFile()):
            e.acceptProposedAction()

    def dropEvent(self, e) -> None:
        path = e.mimeData().urls()[0].toLocalFile()
        if self._is_excel(path):
            self._set_path(path)

    @staticmethod
    def _is_excel(path: str) -> bool:
        return path.lower().endswith((".xlsx", ".xlsm"))

    # -- actions --------------------------------------------------------------
    def _browse(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel Workbook", "",
            "Excel Workbooks (*.xlsx *.xlsm)")
        if path:
            self._set_path(path)

    def _set_path(self, path: str) -> None:
        self._path = path
        self.file_label.setText(os.path.basename(path))
        self.clear_btn.setVisible(True)
        self.status_label.setText("Inspecting workbook…")
        self.progress.setValue(0)
        self._update_applicable_options()
        self._refresh_buttons()

    def _clear_workbook(self) -> None:
        """Remove the selected workbook and reset the form."""
        if self._busy:
            return
        self._path = ""
        self.file_label.setText("No workbook selected.  (You can also drag a file here.)")
        self.clear_btn.setVisible(False)
        self.progress.setValue(0)
        self.status_label.setText("Ready.")
        for cb in self._checkboxes():
            cb.setEnabled(True)
            cb.setChecked(True)
            cb.setToolTip("")
        self._refresh_buttons()

    def _update_applicable_options(self) -> None:
        """Enable only the analyses that make sense for this workbook."""
        try:
            engine = Engine()
            profile = engine.profile(self._path)
            applicable = engine.applicable_options(profile)
        except Exception as exc:                 # noqa: BLE001
            self.status_label.setText(f"⚠  {exc}")
            for cb in self._checkboxes():
                cb.setEnabled(True)
            return
        mapping = {
            "dashboard": self.cb_dashboard, "pivot": self.cb_pivot,
            "kpi": self.cb_kpi, "executive_summary": self.cb_summary,
        }
        n_tables = len(profile.tables)
        for key, cb in mapping.items():
            ok = applicable.get(key, True)
            cb.setEnabled(ok)
            if not ok:
                cb.setChecked(False)
                cb.setToolTip("Not applicable to this workbook's data.")
            else:
                cb.setToolTip("")
        self.status_label.setText(
            f"Ready — detected {n_tables} data table(s). "
            f"Primary: '{profile.primary.sheet_name}' "
            f"({profile.primary.row_count:,} rows)."
        )

    def _selected_options(self) -> AnalysisOptions:
        return AnalysisOptions(
            dashboard=self.cb_dashboard.isChecked(),
            pivot=self.cb_pivot.isChecked(),
            kpi=self.cb_kpi.isChecked(),
            executive_summary=self.cb_summary.isChecked(),
        )

    def _analyze(self) -> None:
        if not self._path:
            QMessageBox.information(self, config.APP_NAME, "Please select a workbook first.")
            return
        options = self._selected_options()
        if not options.any_selected():
            QMessageBox.information(self, config.APP_NAME,
                                    "Please select at least one output type.")
            return
        self._set_busy(True)
        self.status_label.setText("Starting…")
        self._thread, self._worker = start_analysis(
            self._path, options, self._on_progress, self._on_finished, self._on_failed)

    def _on_progress(self, fraction: float, message: str) -> None:
        self.progress.setValue(int(fraction * 100))
        self.status_label.setText(message)

    def _on_finished(self, result) -> None:
        self._set_busy(False)
        self.progress.setValue(100)
        ai = "AI-written" if result.summary_used_llm else "auto-generated"
        msg = (f"Done. Added {len(result.sheets_created)} sheet(s):\n"
               f"  • " + "\n  • ".join(result.sheets_created))
        if any("Executive Summary" in s for s in result.sheets_created):
            msg += f"\n\nExecutive Summary: {ai}."
        for note in result.notes:
            msg += f"\n\n{note}"
        self.status_label.setText("Analysis complete.")
        box = QMessageBox(self)
        box.setWindowTitle(config.APP_NAME)
        box.setText(msg)
        open_btn = box.addButton("Open Workbook", QMessageBox.AcceptRole)
        box.addButton("Close", QMessageBox.RejectRole)
        box.exec()
        if box.clickedButton() is open_btn:
            self._open_workbook()

    def _on_failed(self, error: str) -> None:
        self._set_busy(False)
        self.progress.setValue(0)
        self.status_label.setText("Analysis failed.")
        QMessageBox.critical(self, config.APP_NAME, f"Could not analyze the workbook:\n\n{error}")

    def _open_workbook(self) -> None:
        if self._path and os.path.exists(self._path):
            os.startfile(self._path)  # noqa: S606 - Windows file open

    def _open_folder(self) -> None:
        if self._path and os.path.exists(self._path):
            subprocess.run(["explorer", "/select,", os.path.normpath(self._path)])

    def _open_settings(self) -> None:
        SettingsDialog(self).exec()

    # -- auto-update ----------------------------------------------------------
    def _check_for_updates(self) -> None:
        """Kick off a non-blocking update check on launch (silent if none)."""
        if not config.UPDATE_MANIFEST_URL:
            return
        self.update_label.setText("Checking for updates…")
        self._update_check = UpdateCheckWorker()
        self._update_check.found.connect(self._on_update_found)
        self._update_check.none.connect(lambda: self.update_label.setText("Up to date."))
        self._update_thread = run_on_thread(self._update_check)

    def _on_update_found(self, info) -> None:
        if self._update_thread:
            self._update_thread.quit()
        self.update_label.setText(f"Update available: v{info.version}")
        UpdateDialog(info, self).exec()

    # -- state ----------------------------------------------------------------
    def _checkboxes(self):
        return (self.cb_dashboard, self.cb_pivot, self.cb_kpi, self.cb_summary)

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        self._refresh_buttons()

    def _refresh_buttons(self) -> None:
        has_file = bool(self._path)
        self.analyze_btn.setEnabled(has_file and not self._busy)
        self.browse_btn.setEnabled(not self._busy)
        self.open_wb_btn.setEnabled(has_file and not self._busy)
        self.open_folder_btn.setEnabled(has_file and not self._busy)
        self.clear_btn.setEnabled(not self._busy)
        for cb in self._checkboxes():
            cb.setDisabled(self._busy)
