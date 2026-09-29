"""3. 照合確認：類似度が中間帯の行の対応付け、ひな形外の追加項目。"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QHBoxLayout, QHeaderView, QLabel, QPushButton,
                               QTableWidget, QTabWidget, QVBoxLayout, QWidget)

from core import workflow

from ..widgets import AppState, guarded, ro_item

METHOD_LABELS = {"wbs": "WBS番号", "synonym": "同義語辞書", "exact": "名称一致", "fuzzy": "類似度", "manual": "手動", "none": "-"}
STATUS_LABELS = {"auto": "自動確定", "pending": "確認待ち", "confirmed": "確認済み", "extra": "ひな形外"}
EXTRA = "__extra__"


class MatchingPage(QWidget):
    def __init__(self, state: AppState, parent=None):
        super().__init__(parent)
        self.state = state
        top = QHBoxLayout()
        self.summary = QLabel("")
        top.addWidget(self.summary, 1)
        again = QPushButton("照合をやり直す")
        again.clicked.connect(lambda: self.rematch())
        top.addWidget(again)

        # 確認待ち
        self.pending = QTableWidget(0, 5)
        self.pending.setHorizontalHeaderLabels(["事例", "WBS番号", "作業項目（実用WBS）", "対応するひな形項目", "類似度"])
        self.pending.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.pending.verticalHeader().setVisible(False)
        self.remember = QCheckBox("選択結果を同義語辞書に保存して次回以降も使う")
        self.remember.setChecked(True)
        confirm = QPushButton("選択を確定")
        confirm.setStyleSheet("font-weight:bold")
        confirm.clicked.connect(lambda: self.confirm())
        pend = QWidget()
        pl = QVBoxLayout(pend)
        pl.addWidget(QLabel("類似度が中間帯のため自動で確定しなかった行です。対応するひな形項目を選んでください。"))
        pl.addWidget(self.pending, 1)
        row = QHBoxLayout()
        row.addWidget(self.remember)
        row.addStretch(1)
        row.addWidget(confirm)
        pl.addLayout(row)

        # ひな形外
        self.extras = QTableWidget(0, 4)
        self.extras.setHorizontalHeaderLabels(["事例", "作業項目", "直前のひな形項目", "工程"])
        self.extras.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.extras.verticalHeader().setVisible(False)
        self.extras.setEditTriggers(QAbstractItemView.NoEditTriggers)

        # すべて
        self.all = QTableWidget(0, 6)
        self.all.setHorizontalHeaderLabels(["事例", "WBS番号", "作業項目", "ひな形項目", "方法", "状態"])
        self.all.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.all.verticalHeader().setVisible(False)
        self.all.setEditTriggers(QAbstractItemView.NoEditTriggers)

        tabs = QTabWidget()
        tabs.addTab(pend, "確認待ち")
        tabs.addTab(self.extras, "ひな形外の追加項目")
        tabs.addTab(self.all, "すべての対応付け")
        layout = QVBoxLayout(self)
        layout.addLayout(top)
        layout.addWidget(tabs, 1)
        state.casesChanged.connect(self.refresh)
        state.formatsChanged.connect(self.refresh)
        self.refresh()

    def refresh(self):
        db = self.state.db
        items = {i.key: i for i in workflow.master_items(db)}
        names = {c.id: c.name for c in db.cases()}
        rows = db.case_rows()
        label = lambda k: f"{items[k].wbs_no} {items[k].name}" if k in items else ""
        self.pending.setRowCount(0)
        self.extras.setRowCount(0)
        self.all.setRowCount(0)
        for r in rows:
            if r.status == "pending":
                i = self.pending.rowCount()
                self.pending.insertRow(i)
                self.pending.setItem(i, 0, ro_item(names.get(r.case_id, ""), r.id))
                self.pending.setItem(i, 1, ro_item(r.wbs_no))
                self.pending.setItem(i, 2, ro_item(r.name))
                combo = QComboBox()
                for key, score in r.candidates:
                    if key in items:
                        combo.addItem(f"{label(key)}（{score:.0f}）", key)
                combo.addItem("ひな形外の追加項目として扱う", EXTRA)
                self.pending.setCellWidget(i, 3, combo)
                self.pending.setItem(i, 4, ro_item(f"{r.score:.0f}" if r.score else ""))
            elif r.status == "extra":
                i = self.extras.rowCount()
                self.extras.insertRow(i)
                for col, v in enumerate((names.get(r.case_id, ""), r.name, label(r.prev_key), r.phase)):
                    self.extras.setItem(i, col, ro_item(v))
            i = self.all.rowCount()
            self.all.insertRow(i)
            for col, v in enumerate((names.get(r.case_id, ""), r.wbs_no, r.name, label(r.item_key),
                                     METHOD_LABELS.get(r.method, r.method or ""), STATUS_LABELS.get(r.status, r.status or ""))):
                self.all.setItem(i, col, ro_item(v))
        n_auto = sum(1 for r in rows if r.status in ("auto", "confirmed"))
        self.summary.setText(f"対応付け済み {n_auto}行／確認待ち {self.pending.rowCount()}行／ひな形外 {self.extras.rowCount()}行")

    @guarded("確定できませんでした")
    def confirm(self):
        n = 0
        for i in range(self.pending.rowCount()):
            row_id = self.pending.item(i, 0).data(Qt.UserRole)
            key = self.pending.cellWidget(i, 3).currentData()
            workflow.confirm_match(self.state.db, row_id, None if key == EXTRA else key, self.remember.isChecked())
            n += 1
        self.state.casesChanged.emit()
        self.summary.setText(self.summary.text() + f"（{n}行を確定しました）")

    @guarded("照合できませんでした")
    def rematch(self):
        workflow.rematch(self.state.db, self.state.cfg)
        self.state.casesChanged.emit()
