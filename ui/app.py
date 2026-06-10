"""QApplication bootstrap: theming + main window."""
from __future__ import annotations

import os
import sys

from PySide6.QtGui import QFont, QIcon
from PySide6.QtWidgets import QApplication

import config

from .main_window import MainWindow


def _resource_path(rel: str) -> str:
    """Resolve a bundled resource path (works under PyInstaller onefile)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    # When frozen, resources are bundled at ui/resources; in dev they sit beside this file.
    candidates = [
        os.path.join(base, "ui", "resources", rel),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "resources", rel),
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return candidates[-1]


def _load_stylesheet() -> str:
    path = _resource_path("style.qss")
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return fh.read()
    except OSError:
        return ""


def run() -> int:
    # Remove the leftover old exe from a previous self-update, if any.
    try:
        from core import updater
        updater.cleanup_old()
    except Exception:
        pass

    app = QApplication(sys.argv)
    app.setApplicationName(config.APP_NAME)
    app.setApplicationVersion(config.APP_VERSION)
    app.setFont(QFont("Segoe UI", 10))
    icon_path = _resource_path("app.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    qss = _load_stylesheet()
    if qss:
        app.setStyleSheet(qss)

    window = MainWindow()
    if os.path.exists(icon_path):
        window.setWindowIcon(QIcon(icon_path))
    window.show()
    return app.exec()
