"""1. 様式登録：ひな形・実用WBS様式の登録、プロファイル編集、プレビューでのセル指定。"""
from __future__ import annotations

import copy
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QComboBox, QFileDialog, QFormLayout, QHBoxLayout, QHeaderView, QLabel, QPlainTextEdit,
                               QPushButton, QSpinBox, QSplitter, QTableWidget, QTabWidget, QVBoxLayout, QWidget)

from core import workflow
from core.models import MEANINGS
from excel_io import inspect, profile as profile_mod

from ..widgets import AppState, SheetGrid, guarded, ro_item, run_in_thread

MEANING_ITEMS = [("", "（使わない）")] + list(MEANINGS.items())


class ProfileEditor(QWidget):
    """1 つの様式（kind = master / output）の登録とプロファイル編集。"""

    def __init__(self, state: AppState, kind: str, parent=None):
        super().__init__(parent)
        self.state, self.kind = state, kind
        self.profile: dict | None = None
        self.file: str | None = None
        self.picked: tuple[int, str] | None = None
        title = "ひな形（全作業項目のマスタ）" if kind == "master" else "実用WBS様式（出力テンプレート）"

        top = QHBoxLayout()
        self.register_btn = QPushButton(f"{title}を登録…")
        self.register_btn.clicked.connect(lambda: self.choose_file())
        top.addWidget(self.register_btn)
        self.file_label = QLabel("未登録")
        top.addWidget(self.file_label, 1)
        self.save_btn = QPushButton("プロファイルを保存")
        self.save_btn.setStyleSheet("font-weight:bold")
        self.save_btn.clicked.connect(lambda: self.save_profile())
        top.addWidget(self.save_btn)

        self.warning = QLabel("")
        self.warning.setWordWrap(True)
        self.warning.setStyleSheet("background:#fff3cd; color:#8a5a00; padding:4px")
        self.warning.hide()

        form = QFormLayout()
        self.sheet = QComboBox()
        self.sheet.currentTextChanged.connect(lambda s: self._sheet_changed(s))
        self.header_row = QSpinBox()
        self.data_start = QSpinBox()
        self.data_end = QSpinBox()
        for sp in (self.header_row, self.data_start, self.data_end):
            sp.setRange(0, 100000)
            sp.setSpecialValueText("（未指定）")
            sp.valueChanged.connect(lambda _v: self._spins_changed())
        form.addRow("対象シート", self.sheet)
        form.addRow("ヘッダー行", self.header_row)
        form.addRow("データ開始行", self.data_start)
        if kind == "output":
            form.addRow("データ終了行（明細の最終行）", self.data_end)
        self.conf = QLabel("")
        form.addRow("確信度", self.conf)

        self.cols = QTableWidget(0, 3)
        self.cols.setHorizontalHeaderLabels(["列", "見出し", "意味"])
        self.cols.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.cols.verticalHeader().setVisible(False)

        pick = QHBoxLayout()
        self.pick_label = QLabel("プレビューのセルをクリックして指定できます")
        self.pick_label.setStyleSheet("color:#555")
        pick.addWidget(self.pick_label, 1)
        for text, what in (("この行をヘッダー行に", "header"), ("この行をデータ開始行に", "start")) + (
                (("この行をデータ終了行に", "end"),) if kind == "output" else ()):
            b = QPushButton(text)
            b.clicked.connect(lambda _=False, w=what: self.apply_pick(w))
            pick.addWidget(b)
        self.meaning_pick = QComboBox()
        for k, label in MEANING_ITEMS:
            self.meaning_pick.addItem(label, k)
        pick.addWidget(self.meaning_pick)
        b = QPushButton("この列の意味にする")
        b.clicked.connect(lambda: self.apply_pick("column"))
        pick.addWidget(b)

        self.analysis = QPlainTextEdit()
        self.analysis.setReadOnly(True)
        self.analysis.setMaximumHeight(170)

        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 0, 0, 0)
        ll.addLayout(form)
        ll.addWidget(QLabel("各列の意味"))
        ll.addWidget(self.cols, 1)
        if kind == "output":
            ll.addWidget(QLabel("様式の解析結果（書き込み時に保護するもの）"))
            ll.addWidget(self.analysis)
        self.grid = SheetGrid()
        self.grid.cellPicked.connect(self._picked)
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.addWidget(QLabel("プレビュー（青＝ヘッダー行、黄＝データ範囲）"))
        rl.addWidget(self.grid, 1)
        rl.addLayout(pick)
        split = QSplitter(Qt.Horizontal)
        split.addWidget(left)
        split.addWidget(right)
        split.setSizes([420, 900])

        layout = QVBoxLayout(self)
        layout.addLayout(top)
        layout.addWidget(self.warning)
        layout.addWidget(split, 1)
        self.load_from_db()

    # ---- 表示 ----
    def load_from_db(self):
        f = self.state.db.active_format(self.kind)
        if not f:
            return
        self.file = f["file"]
        self.profile = copy.deepcopy(f["profile"])
        self.file_label.setText(f"{Path(f['source']).name}（登録日 {f['created'][:10]}）")
        self._fill(f.get("analysis"))

    def _fill(self, analysis=None):
        p = self.profile or {}
        self._filling = True
        self.sheet.blockSignals(True)
        self.sheet.clear()
        self.sheet.addItems(p.get("sheets") or [p.get("sheet", "")])
        self.sheet.setCurrentText(p.get("sheet", ""))
        self.sheet.blockSignals(False)
        for sp, key in ((self.header_row, "header_row"), (self.data_start, "data_start"), (self.data_end, "data_end")):
            sp.blockSignals(True)
            sp.setValue(int(p.get(key) or 0))
            sp.blockSignals(False)
        self.conf.setText(f"{p.get('confidence', 0):.2f}")
        th = float(self.state.cfg.get("profile", {}).get("confidence_threshold", 0.7))
        issues = p.get("issues") or []
        if p and (p.get("confidence", 0) < th or issues):
            self.warning.setText("自動推定に自信がありません。プレビューのセルをクリックして、ヘッダー行・データ行・各列の意味を指定してください。\n"
                                 + "\n".join("・" + i for i in issues))
            self.warning.show()
        else:
            self.warning.hide()
        self._fill_cols()
        if self.file:
            self.grid.load(self.file, p.get("sheet", ""))
            self.grid.highlight(p)
        if analysis and self.kind == "output":
            self.analysis.setPlainText("\n".join(inspect.summary_lines(analysis)))
        self._filling = False

    def _fill_cols(self):
        p = self.profile or {}
        self.cols.setRowCount(0)
        for col, meaning in sorted(p.get("columns", {}).items(), key=lambda kv: (len(kv[0]), kv[0])):
            r = self.cols.rowCount()
            self.cols.insertRow(r)
            self.cols.setItem(r, 0, ro_item(col))
            self.cols.setItem(r, 1, ro_item(p.get("headers", {}).get(col, "")))
            combo = QComboBox()
            for k, label in MEANING_ITEMS:
                combo.addItem(label, k)
            combo.setCurrentIndex(max(0, combo.findData(meaning)))
            combo.currentIndexChanged.connect(lambda _i, c=col, w=combo: self._set_meaning(c, w.currentData()))
            self.cols.setCellWidget(r, 2, combo)

    # ---- 編集 ----
    def _sheet_changed(self, sheet):
        if getattr(self, "_filling", False) or not self.profile or not sheet:
            return
        self.profile["sheet"] = sheet
        self.grid.load(self.file, sheet)
        self.grid.highlight(self.profile)

    def _spins_changed(self):
        if getattr(self, "_filling", False) or not self.profile:
            return
        self.profile["header_row"] = self.header_row.value() or None
        self.profile["data_start"] = self.data_start.value() or None
        if self.kind == "output":
            self.profile["data_end"] = self.data_end.value() or None
        self._refresh_headers()
        self.grid.highlight(self.profile)

    def _refresh_headers(self):
        hr = self.profile.get("header_row")
        if not hr:
            return
        headers = {}
        for col in self.profile.get("columns", {}):
            item = self.grid.item(hr - 1, SheetGrid._col_index(col) - 1)
            headers[col] = item.text() if item else ""
        self.profile["headers"] = headers
        self._fill_cols()

    def _set_meaning(self, col, meaning):
        if not self.profile:
            return
        if meaning:
            self.profile["columns"][col] = meaning
        else:
            self.profile["columns"].pop(col, None)
        self._fill_cols()
        self.grid.highlight(self.profile)

    def _picked(self, row, col):
        self.picked = (row, col)
        item = self.grid.item(row - 1, SheetGrid._col_index(col) - 1)
        self.pick_label.setText(f"選択：{col}{row}「{item.text() if item else ''}」")

    @guarded("指定できませんでした")
    def apply_pick(self, what: str):
        if not self.profile or not self.picked:
            return
        row, col = self.picked
        if what == "header":
            self.header_row.setValue(row)
            if not self.profile.get("data_start") or self.profile["data_start"] <= row:
                self.data_start.setValue(row + 1)
        elif what == "start":
            self.data_start.setValue(row)
        elif what == "end":
            self.data_end.setValue(row)
        elif what == "column":
            meaning = self.meaning_pick.currentData()
            for c, m in list(self.profile["columns"].items()):
                if meaning and m == meaning:
                    del self.profile["columns"][c]
            self._set_meaning(col, meaning)
            self._refresh_headers()
        self.profile["issues"] = []
        self.profile["confidence"] = 1.0
        self.profile["manual"] = True
        self.conf.setText("1.00（手動で指定）")
        self.warning.hide()

    # ---- 登録・保存 ----
    @guarded("登録できませんでした")
    def choose_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "様式を選択", "", "Excel ブック (*.xlsx *.xlsm)")
        if path:
            self.register(path)

    def register(self, path):
        fn = workflow.register_master if self.kind == "master" else workflow.register_output
        result = fn(self.state.db, path, self.state.cfg)
        self.load_from_db()
        self.state.formatsChanged.emit()
        return result

    @guarded("保存できませんでした")
    def save_profile(self):
        if not self.profile:
            return
        if not any(m == "name" for m in self.profile.get("columns", {}).values()):
            raise ValueError("「作業項目名」の列を指定してください。")
        if self.kind == "master":
            n = workflow.reload_master(self.state.db, self.state.cfg, self.profile)
            self.pick_label.setText(f"保存しました（ひな形 {n}項目）")
        else:
            analysis = workflow.update_output_profile(self.state.db, self.profile)
            self.analysis.setPlainText("\n".join(inspect.summary_lines(analysis)))
            self.pick_label.setText("保存しました")
        self.state.formatsChanged.emit()


class FormatPage(QWidget):
    def __init__(self, state: AppState, parent=None):
        super().__init__(parent)
        self.master = ProfileEditor(state, "master")
        self.output = ProfileEditor(state, "output")
        tabs = QTabWidget()
        tabs.addTab(self.master, "ひな形")
        tabs.addTab(self.output, "実用WBS様式（出力テンプレート）")
        layout = QVBoxLayout(self)
        layout.addWidget(tabs)
