"""Custom Generate wizard: pick the titles to study and the values to total."""
from __future__ import annotations

import config
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QDialog,
                               QFrame, QHBoxLayout, QLabel, QListWidget,
                               QListWidgetItem, QMessageBox, QPushButton,
                               QScrollArea, QVBoxLayout, QWidget)

from core.models import CustomSelection, MeasureChoice

# (display label, format_kind) shown in the per-value dropdown.
_FORMATS = [
    ("USD ($)", "usd"),
    ("Lebanese Pound (LBP)", "lbp"),
    ("Number", "number"),
    ("Percent (%)", "percent"),
    ("Auto-detect", "auto"),
]
_UNIT_DEFAULT = {"currency": "usd", "percent": "percent", "number": "number"}


class CustomWizardDialog(QDialog):
    def __init__(self, engine, profile, parent=None) -> None:
        super().__init__(parent)
        self._engine = engine
        self._profile = profile
        self._dim_checks: list[tuple[QCheckBox, str]] = []
        self._measure_rows: list[tuple[QCheckBox, str, QComboBox]] = []
        self.selection: CustomSelection | None = None

        self.setWindowTitle("Custom Generate")
        self.setMinimumSize(580, 740)
        self._desc = engine.describe_columns(profile)
        self._build()

    # -- layout ---------------------------------------------------------------
    def _build(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(22, 20, 22, 18)
        root.setSpacing(12)

        title = QLabel("Custom Generate")
        title.setObjectName("TitleLabel")
        root.addWidget(title)

        # Optional sheet picker (only when several data sheets exist).
        if len(self._desc.get("sheets", [])) > 1:
            row = QHBoxLayout()
            row.addWidget(QLabel("Data sheet:"))
            self.sheet_combo = QComboBox()
            self.sheet_combo.addItems(self._desc["sheets"])
            if self._desc.get("sheet"):
                i = self.sheet_combo.findText(self._desc["sheet"])
                if i >= 0:
                    self.sheet_combo.setCurrentIndex(i)
            self.sheet_combo.currentTextChanged.connect(self._on_sheet_changed)
            row.addWidget(self.sheet_combo, 1)
            root.addLayout(row)
        else:
            self.sheet_combo = None

        root.addWidget(self._section("1.  Select the titles to study (group by)"))
        self._dims_host = self._scroll_area()
        root.addWidget(self._dims_host["scroll"], 1)

        root.addWidget(self._section("2.  Select the values to analyze (and format)"))
        self._meas_host = self._scroll_area()
        root.addWidget(self._meas_host["scroll"], 1)

        self._populate()

        # --- Combination pivots section (optional, in addition to singles) ---
        root.addWidget(self._section(
            "3.  Combination pivots (optional) — nest titles into one table"))
        combo_row = QHBoxLayout()
        self.combo_src = QListWidget()
        self.combo_src.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.combo_src.setMaximumHeight(96)
        combo_row.addWidget(self.combo_src, 1)
        btns = QVBoxLayout()
        add_btn = QPushButton("Add ▸")
        add_btn.setToolTip("Pick 2–3 titles on the left, then Add a combination")
        add_btn.clicked.connect(self._add_combination)
        rm_btn = QPushButton("Remove")
        rm_btn.clicked.connect(self._remove_combination)
        btns.addWidget(add_btn)
        btns.addWidget(rm_btn)
        btns.addStretch(1)
        combo_row.addLayout(btns)
        self.combo_added = QListWidget()
        self.combo_added.setMaximumHeight(96)
        combo_row.addWidget(self.combo_added, 1)
        root.addLayout(combo_row)
        self._populate_combos()

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        cancel = QPushButton("Cancel")
        cancel.clicked.connect(self.reject)
        gen = QPushButton("Generate study")
        gen.setObjectName("PrimaryButton")
        gen.clicked.connect(self._on_generate)
        buttons.addWidget(cancel)
        buttons.addWidget(gen)
        root.addLayout(buttons)

    def _section(self, text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setObjectName("SectionLabel")
        return lbl

    def _scroll_area(self) -> dict:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        inner = QWidget()
        layout = QVBoxLayout(inner)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(6)
        layout.addStretch(1)
        scroll.setWidget(inner)
        return {"scroll": scroll, "inner": inner, "layout": layout}

    # -- population -----------------------------------------------------------
    def _clear(self, host: dict) -> None:
        layout = host["layout"]
        while layout.count() > 1:                 # keep the trailing stretch
            item = layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()

    def _populate(self) -> None:
        self._dim_checks.clear()
        self._measure_rows.clear()
        self._clear(self._dims_host)
        self._clear(self._meas_host)

        for d in self._desc["dimensions"]:
            row = QWidget()
            h = QHBoxLayout(row)
            h.setContentsMargins(0, 0, 0, 0)
            cb = QCheckBox(d["name"])
            cb.setChecked(bool(d.get("recommended")))
            h.addWidget(cb)
            badge = QLabel(f"{d['kind']} · {d['detail']}")
            badge.setObjectName("HintLabel")
            h.addWidget(badge)
            h.addStretch(1)
            self._dims_host["layout"].insertWidget(self._dims_host["layout"].count() - 1, row)
            self._dim_checks.append((cb, d["name"]))

        for m in self._desc["measures"]:
            row = QWidget()
            h = QHBoxLayout(row)
            h.setContentsMargins(0, 0, 0, 0)
            cb = QCheckBox(m["name"])
            cb.setChecked(bool(m.get("recommended")))
            h.addWidget(cb)
            h.addStretch(1)
            combo = QComboBox()
            for label, _kind in _FORMATS:
                combo.addItem(label)
            default_kind = _UNIT_DEFAULT.get(m.get("unit", "number"), "number")
            for idx, (_label, kind) in enumerate(_FORMATS):
                if kind == default_kind:
                    combo.setCurrentIndex(idx)
                    break
            combo.setFixedWidth(190)
            h.addWidget(combo)
            self._meas_host["layout"].insertWidget(self._meas_host["layout"].count() - 1, row)
            self._measure_rows.append((cb, m["name"], combo))

    def _populate_combos(self) -> None:
        self.combo_src.clear()
        self.combo_added.clear()
        for d in self._desc["dimensions"]:
            self.combo_src.addItem(d["name"])

    def _add_combination(self) -> None:
        names = [i.text() for i in self.combo_src.selectedItems()]
        if len(names) < 2:
            QMessageBox.information(self, config.APP_NAME,
                                    "Pick at least 2 titles (Ctrl-click) to combine.")
            return
        item = QListWidgetItem(" + ".join(names))
        item.setData(Qt.UserRole, names)
        self.combo_added.addItem(item)
        self.combo_src.clearSelection()

    def _remove_combination(self) -> None:
        for it in self.combo_added.selectedItems():
            self.combo_added.takeItem(self.combo_added.row(it))

    def _on_sheet_changed(self, sheet: str) -> None:
        self._desc = self._engine.describe_columns(self._profile, sheet)
        self._populate()
        self._populate_combos()

    # -- result ---------------------------------------------------------------
    def _on_generate(self) -> None:
        dims = [name for cb, name in self._dim_checks if cb.isChecked()]
        measures = []
        for cb, name, combo in self._measure_rows:
            if cb.isChecked():
                kind = _FORMATS[combo.currentIndex()][1]
                measures.append(MeasureChoice(name, kind))
        if not measures:
            QMessageBox.information(self, config.APP_NAME,
                                    "Please select at least one value to analyze.")
            return
        combinations = [self.combo_added.item(i).data(Qt.UserRole)
                        for i in range(self.combo_added.count())]
        sheet = self.sheet_combo.currentText() if self.sheet_combo else self._desc.get("sheet")
        self.selection = CustomSelection(sheet_name=sheet, dimensions=dims,
                                         measures=measures,
                                         combinations=combinations)
        self.accept()
