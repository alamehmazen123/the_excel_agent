"""Settings dialog: Groq API key + model. Key saved to Credential Manager."""
from __future__ import annotations

import config
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QComboBox, QDialog, QFormLayout, QHBoxLayout,
                               QLabel, QLineEdit, QMessageBox, QPushButton,
                               QVBoxLayout)

from core.llm.groq_client import GroqNarrator

from . import settings_store


class SettingsDialog(QDialog):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setMinimumWidth(460)
        self._build()

    def _build(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 20)
        layout.setSpacing(14)

        title = QLabel("AI Narrative Settings")
        title.setObjectName("TitleLabel")
        layout.addWidget(title)

        info = QLabel(
            "The Executive Summary can be written by Groq's free AI. Paste your "
            "own free key (groq.com) for best reliability. If left blank, a "
            "built-in default is used; if AI is unavailable, summaries are still "
            "generated from your data automatically."
        )
        info.setObjectName("SubtitleLabel")
        info.setWordWrap(True)
        layout.addWidget(info)

        warn = QLabel("⚠  DO NOT CHANGE, unless you are aware of what you are "
                      "doing (have a new key).")
        warn.setObjectName("WarnLabel")
        warn.setWordWrap(True)
        layout.addWidget(warn)

        form = QFormLayout()
        form.setSpacing(10)

        self.key_edit = QLineEdit()
        self.key_edit.setEchoMode(QLineEdit.Password)
        self.key_edit.setPlaceholderText("gsk_…")
        self.key_edit.setText(settings_store.get_user_key())
        form.addRow("Groq API key:", self.key_edit)

        self.model_combo = QComboBox()
        self.model_combo.addItems(config.AVAILABLE_GROQ_MODELS)
        current = settings_store.get_model()
        idx = self.model_combo.findText(current)
        if idx >= 0:
            self.model_combo.setCurrentIndex(idx)
        form.addRow("Model:", self.model_combo)
        layout.addLayout(form)

        self.status = QLabel("")
        self.status.setObjectName("StatusLabel")
        self.status.setWordWrap(True)
        layout.addWidget(self.status)

        buttons = QHBoxLayout()
        test_btn = QPushButton("Test Connection")
        test_btn.clicked.connect(self._test)
        buttons.addWidget(test_btn)
        buttons.addStretch(1)
        cancel = QPushButton("Cancel")
        cancel.clicked.connect(self.reject)
        save = QPushButton("Save")
        save.setObjectName("PrimaryButton")
        save.clicked.connect(self._save)
        buttons.addWidget(cancel)
        buttons.addWidget(save)
        layout.addLayout(buttons)

    def _current_effective_key(self) -> str:
        return self.key_edit.text().strip() or config.bundled_groq_key()

    def _test(self) -> None:
        key = self._current_effective_key()
        if not key:
            self.status.setText("No key entered and no built-in default available.")
            return
        self.status.setText("Testing…")
        self.repaint()
        ok, msg = GroqNarrator(key, model=self.model_combo.currentText()).test_connection()
        self.status.setText(("✓ " if ok else "✗ ") + msg)

    def _save(self) -> None:
        settings_store.set_user_key(self.key_edit.text().strip())
        settings_store.set_model(self.model_combo.currentText())
        self.accept()
