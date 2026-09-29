"""2. 過去事例管理：実用WBSの一括取込（ドラッグ＆ドロップ対応）・属性編集・一覧・削除。"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QAbstractItemView, QFileDialog, QHBoxLayout, QHeaderView, QLabel, QLineEdit,
                               QMessageBox, QPlainTextEdit, QPushButton, QSplitter, QTableWidget, QVBoxLayout, QWidget)

from core import workflow

from ..widgets import AppState, AttrsEditor, guarded, ro_item, run_in_thread


class DropArea(QLabel):
    def __init__(self, on_files, parent=None):
        super().__init__("ここに実用WBS（.xlsx / .xlsm）をドラッグ＆ドロップ（複数可）", parent)
        self.on_files = on_files
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignCenter)
        self.setMinimumHeight(60)
        self.setStyleSheet("border:2px dashed #8aa; color:#557; background:#f7fafc")

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        files = [u.toLocalFile() for u in e.mimeData().urls() if u.toLocalFile()]
        self.on_files(files)


class CasesPage(QWidget):
    def __init__(self, state: AppState, parent=None):
        super().__init__(parent)
        self.state = state
        top = QHBoxLayout()
        add = QPushButton("ファイルを追加…")
        add.clicked.connect(lambda: self.choose_files())
        top.addWidget(add)
        rm = QPushButton("選んだ事例を削除")
        rm.clicked.connect(lambda: self.delete_selected())
        top.addWidget(rm)
        top.addStretch(1)
        self.count = QLabel("")
        top.addWidget(self.count)

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["事例名", "ファイル", "取込日時", "行数", "案件属性"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.table.itemSelectionChanged.connect(self._selected)

        edit = QWidget()
        el = QVBoxLayout(edit)
        el.addWidget(QLabel("案件属性（任意。未入力でも判定できます）"))
        self.name = QLineEdit()
        el.addWidget(self.name)
        self.attrs = AttrsEditor(state.cfg)
        el.addWidget(self.attrs)
        save = QPushButton("属性を保存")
        save.clicked.connect(lambda: self.save_attrs())
        el.addWidget(save)
        el.addStretch(1)

        split = QSplitter(Qt.Horizontal)
        split.addWidget(self.table)
        split.addWidget(edit)
        split.setSizes([900, 420])
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(120)

        layout = QVBoxLayout(self)
        layout.addWidget(DropArea(self.import_files))
        layout.addLayout(top)
        layout.addWidget(split, 1)
        layout.addWidget(QLabel("取込ログ"))
        layout.addWidget(self.log)
        state.casesChanged.connect(self.refresh)
        state.formatsChanged.connect(self.refresh)
        self.refresh()

    def refresh(self):
        db = self.state.db
        rows = db.case_rows()
        counts = {}
        for r in rows:
            counts[r.case_id] = counts.get(r.case_id, 0) + 1
        self.table.setRowCount(0)
        for c in db.cases():
            i = self.table.rowCount()
            self.table.insertRow(i)
            a = c.attrs or {}
            desc = "／".join(str(x) for x in (a.get("update_type"), f"{a['beds']}床" if a.get("beds") else None,
                                             a.get("vendor"), "・".join(a.get("systems") or []) or None,
                                             f"{a['year']}年度" if a.get("year") else None) if x)
            for col, v in enumerate((c.name, Path(c.source).name, c.imported.replace("T", " "), counts.get(c.id, 0),
                                     desc or "（未入力）")):
                self.table.setItem(i, col, ro_item(v, c.id))
        self.count.setText(f"{self.table.rowCount()}件")

    def _selected_id(self):
        rows = {i.row() for i in self.table.selectedIndexes()}
        return self.table.item(min(rows), 0).data(Qt.UserRole) if rows else None

    def _selected(self):
        cid = self._selected_id()
        case = next((c for c in self.state.db.cases() if c.id == cid), None)
        if case:
            self.name.setText(case.name)
            self.attrs.set_attrs(case.attrs)

    @guarded("保存できませんでした")
    def save_attrs(self):
        cid = self._selected_id()
        if cid is None:
            return
        self.state.db.update_case_attrs(cid, self.attrs.attrs())
        if self.name.text().strip():
            self.state.db.rename_case(cid, self.name.text().strip())
        self.state.casesChanged.emit()

    @guarded("削除できませんでした")
    def delete_selected(self):
        ids = {self.table.item(i.row(), 0).data(Qt.UserRole) for i in self.table.selectedIndexes()}
        if not ids or QMessageBox.question(self, "削除の確認", f"{len(ids)}件の事例を削除しますか？") != QMessageBox.Yes:
            return
        for cid in ids:
            self.state.db.delete_case(cid)
        self.state.casesChanged.emit()

    @guarded("取り込めませんでした")
    def choose_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "実用WBSを選択", "", "Excel ブック (*.xlsx *.xlsm)")
        if files:
            self.import_files(files)

    def import_files(self, files, sync: bool = False):
        if not self.state.db.active_format("master"):
            QMessageBox.information(self, "過去事例", "先に「様式登録」でひな形を登録してください。")
            return
        state = self.state

        def work(progress):
            log, dups = [], []
            for n, f in enumerate(files):
                progress(n * 100 / max(1, len(files)), Path(f).name)
                try:
                    workflow.import_case(state.db, f, state.cfg)
                    log.append(f"取込：{Path(f).name}")
                except workflow.DuplicateCase as e:
                    dups.append(f)
                    log.append(f"警告：{e}")
                except Exception as e:
                    from ..widgets import error_message
                    log.append(f"失敗：{Path(f).name}（{error_message(e)}）")
            return log, dups

        def done(result):
            log, dups = result
            for l in log:
                self.log.appendPlainText(l)
            if dups:
                names = "\n".join(Path(d).name for d in dups)
                ans = QMessageBox.warning(self, "二重取込", f"次のファイルは取込済みです（内容が同じファイル）。\n{names}\n\nそれでも取り込みますか？",
                                          QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if ans == QMessageBox.Yes:
                    for d in dups:
                        workflow.import_case(state.db, d, state.cfg, allow_duplicate=True)
                        self.log.appendPlainText(f"取込（二重）：{Path(d).name}")
            state.casesChanged.emit()

        if sync:
            done(work(lambda *_: None))
        else:
            run_in_thread(self, work, done, "過去事例の取込")
