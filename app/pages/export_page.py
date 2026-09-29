"""6. Excel出力：書き込みエンジンの表示、列の対応、差分プレビュー、セル選択、保存、検証結果。"""
from __future__ import annotations

import os
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (QComboBox, QHBoxLayout, QHeaderView, QLabel, QMessageBox, QPlainTextEdit, QPushButton,
                               QSplitter, QTableWidget, QTabWidget, QVBoxLayout, QWidget)

from core import paths, workflow
from core.models import MEANINGS
from excel_io import com_writer, engine as engine_mod, plan as plan_mod

from ..widgets import AppState, SheetGrid, guarded, ro_item, run_in_thread

SOURCE_ITEMS = [("", "（書き込まない）"), ("__seq__", "連番（1, 2, 3 …）")] + list(MEANINGS.items())
STATUS_LABELS = {"write": "書き込む", "clear": "値を消す", "skip": "書き込まない"}


class ExportPage(QWidget):
    def __init__(self, state: AppState, parent=None):
        super().__init__(parent)
        self.state = state
        self.plan = None
        self.sources: dict = {}
        self.picked = None
        self.last_output: Path | None = None

        top = QHBoxLayout()
        self.engine_label = QLabel("")
        top.addWidget(self.engine_label, 1)
        self.engine = QComboBox()
        self.engine.addItem("自動（Excel があれば COM 方式）", "auto")
        self.engine.addItem("Excel COM 方式", "com")
        self.engine.addItem("OOXML 直接編集方式", "ooxml")
        top.addWidget(QLabel("書き込み方式"))
        top.addWidget(self.engine)
        prev = QPushButton("差分プレビューを作成")
        prev.clicked.connect(lambda: self.make_plan())
        top.addWidget(prev)
        self.run_btn = QPushButton("確認して出力")
        self.run_btn.setStyleSheet("font-weight:bold")
        self.run_btn.clicked.connect(lambda: self.export())
        self.run_btn.setEnabled(False)
        top.addWidget(self.run_btn)

        self.mapping = QTableWidget(0, 3)
        self.mapping.setHorizontalHeaderLabels(["テンプレートの列", "見出し", "書き込む値（ひな形の列）"])
        self.mapping.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.mapping.verticalHeader().setVisible(False)
        save_map = QPushButton("列の対応を様式プロファイルに保存")
        save_map.clicked.connect(lambda: self.save_sources())
        map_box = QWidget()
        mb = QVBoxLayout(map_box)
        mb.setContentsMargins(0, 0, 0, 0)
        mb.addWidget(QLabel("テンプレートの各列に、ひな形のどの値を書くか（判断できない列は選んでください）"))
        mb.addWidget(self.mapping, 1)
        mb.addWidget(save_map)

        self.diff = QTableWidget(0, 5)
        self.diff.setHorizontalHeaderLabels(["セル", "現在の値", "書き込む値", "処理", "理由"])
        self.diff.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.diff.verticalHeader().setVisible(False)
        self.warnings = QPlainTextEdit()
        self.warnings.setReadOnly(True)
        self.warnings.setMaximumHeight(110)
        diff_box = QWidget()
        db_ = QVBoxLayout(diff_box)
        db_.setContentsMargins(0, 0, 0, 0)
        self.plan_label = QLabel("［差分プレビューを作成］を押すと、書き込む内容を確認できます。")
        self.plan_label.setWordWrap(True)
        db_.addWidget(self.plan_label)
        db_.addWidget(self.diff, 1)
        db_.addWidget(QLabel("警告"))
        db_.addWidget(self.warnings)

        self.grid = SheetGrid()
        self.grid.cellPicked.connect(self._picked)
        grid_box = QWidget()
        gb = QVBoxLayout(grid_box)
        gb.setContentsMargins(0, 0, 0, 0)
        gb.addWidget(QLabel("テンプレート（オレンジ＝書き込むセル）。書き込み先に迷う場合はセルをクリックして指定します。"))
        gb.addWidget(self.grid, 1)
        pick = QHBoxLayout()
        self.pick_label = QLabel("")
        pick.addWidget(self.pick_label, 1)
        for text, what in (("この行をデータ開始行に", "start"), ("この行を明細の最終行に", "end")):
            b = QPushButton(text)
            b.clicked.connect(lambda _=False, w=what: self.apply_pick(w))
            pick.addWidget(b)
        gb.addLayout(pick)

        self.result = QTableWidget(0, 3)
        self.result.setHorizontalHeaderLabels(["検証項目", "結果", "詳細"])
        self.result.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.result.verticalHeader().setVisible(False)
        self.result_label = QLabel("")
        open_dir = QPushButton("出力フォルダを開く")
        open_dir.clicked.connect(lambda: os.startfile(str(self.last_output.parent if self.last_output else paths.output_dir())))
        res_box = QWidget()
        rb = QVBoxLayout(res_box)
        rb.setContentsMargins(0, 0, 0, 0)
        row = QHBoxLayout()
        row.addWidget(self.result_label, 1)
        row.addWidget(open_dir)
        rb.addLayout(row)
        rb.addWidget(self.result, 1)

        tabs = QTabWidget()
        tabs.addTab(diff_box, "差分プレビュー")
        tabs.addTab(grid_box, "テンプレートのプレビュー")
        tabs.addTab(res_box, "検証結果")
        self.tabs = tabs
        split = QSplitter(Qt.Horizontal)
        split.addWidget(map_box)
        split.addWidget(tabs)
        split.setSizes([420, 900])
        layout = QVBoxLayout(self)
        layout.addLayout(top)
        layout.addWidget(split, 1)
        state.formatsChanged.connect(self.refresh)
        state.settingsChanged.connect(self.refresh)
        self.refresh()

    # ---- 表示 ----
    def refresh(self):
        excel = com_writer.excel_available()
        chosen = engine_mod.choose(self.state.cfg.get("output", {}).get("engine", "auto"))
        self.engine.setCurrentIndex(max(0, self.engine.findData(self.state.cfg.get("output", {}).get("engine", "auto"))))
        self.engine_label.setText(f"Excel：{'インストールあり' if excel else 'なし'}　→　既定の方式：{engine_mod.ENGINE_LABELS[chosen]}")
        out = self.state.db.active_format("output")
        master = self.state.db.active_format("master")
        self.mapping.setRowCount(0)
        if not out:
            self.plan_label.setText("先に「様式登録」で実用WBS様式（出力テンプレート）を登録してください。")
            return
        prof = out["profile"]
        meanings = set((master or {}).get("profile", {}).get("columns", {}).values())
        self.sources = dict(prof.get("sources") or plan_mod.default_sources(prof, meanings, out["analysis"]))
        for col in sorted(prof.get("columns", {}), key=lambda c: (len(c), c)):
            i = self.mapping.rowCount()
            self.mapping.insertRow(i)
            self.mapping.setItem(i, 0, ro_item(f"{col}列（{MEANINGS.get(prof['columns'][col], '')}）", col))
            self.mapping.setItem(i, 1, ro_item(prof.get("headers", {}).get(col, "")))
            combo = QComboBox()
            for k, label in SOURCE_ITEMS:
                combo.addItem(label, k)
            combo.setCurrentIndex(max(0, combo.findData(self.sources.get(col) or "")))
            if not self.sources.get(col):
                combo.setStyleSheet("background:#fff3cd")
            combo.currentIndexChanged.connect(lambda _i, c=col, w=combo: self.sources.__setitem__(c, w.currentData() or None))
            self.mapping.setCellWidget(i, 2, combo)
        try:
            self.grid.load(out["file"], prof["sheet"])
            self.grid.highlight(prof)
        except Exception:
            pass

    @guarded("保存できませんでした")
    def save_sources(self):
        out = self.state.db.active_format("output")
        prof = dict(out["profile"])
        prof["sources"] = self.sources
        self.state.db.update_format_profile(out["id"], prof)
        self.plan_label.setText("列の対応を保存しました。")

    def _picked(self, row, col):
        self.picked = (row, col)
        self.pick_label.setText(f"選択：{col}{row}")

    @guarded("指定できませんでした")
    def apply_pick(self, what):
        if not self.picked:
            return
        out = self.state.db.active_format("output")
        prof = dict(out["profile"])
        prof["data_start" if what == "start" else "data_end"] = self.picked[0]
        prof["sources"] = self.sources
        workflow.update_output_profile(self.state.db, prof)
        self.state.formatsChanged.emit()
        self.pick_label.setText(f"{'データ開始行' if what == 'start' else '明細の最終行'}を {self.picked[0]} 行目にしました（様式プロファイルに保存）")

    # ---- 計画 ----
    @guarded("差分プレビューを作成できませんでした")
    def make_plan(self):
        if not self.state.judgments:
            raise ValueError("先に「判定結果」で判定を実行し、出力する項目を確定してください。")
        setting = self.engine.currentData()
        self.state.cfg.setdefault("output", {})["engine"] = setting
        eng = engine_mod.choose(setting)
        self.plan = workflow.build_plan(self.state.db, self.state.cfg, self.state.judgments, self.sources, eng)
        p = self.plan
        self.diff.setRowCount(0)
        for w in p.writes:
            i = self.diff.rowCount()
            self.diff.insertRow(i)
            for col, v in enumerate((w.ref, w.current, "" if w.value is None else plan_mod.text_of(w.value),
                                     STATUS_LABELS[w.status], w.reason)):
                item = ro_item(v)
                if w.status == "skip":
                    item.setForeground(QBrush(QColor("#b35c00")))
                self.diff.setItem(i, col, item)
        head = (f"{engine_mod.ENGINE_LABELS[p.engine]}：出力 {p.rows_needed}行（テンプレートの明細 {p.capacity}行"
                + (f"、{p.insert_rows}行を挿入" if p.insert_rows else "") + f"）、書き込み {len(p.actions)}セル、書かないセル {len(p.skipped)}")
        if p.blocked:
            head = "書き込めません：" + p.blocked
        self.plan_label.setText(head)
        self.plan_label.setStyleSheet("color:#b00020" if p.blocked else "")
        self.warnings.setPlainText("\n".join(p.warnings) or "（なし）")
        self.grid.highlight(self.state.db.active_format("output")["profile"], {w.ref for w in p.actions})
        self.run_btn.setEnabled(not p.blocked)
        self.tabs.setCurrentIndex(0)
        return p

    # ---- 出力 ----
    def export(self, confirm: bool = True, out_dir: Path | None = None, sync: bool = False):
        if not self.plan or self.plan.blocked:
            return None
        if confirm and QMessageBox.question(
                self, "出力の確認", f"差分プレビューのとおり {len(self.plan.actions)}セルを書き込み、"
                                   f"テンプレートのコピーとして新しいファイルを作ります。よろしいですか？") != QMessageBox.Yes:
            return None
        st, plan = self.state, self.plan

        def work(progress):
            progress(20, "書き込んでいます")
            return workflow.export(st.db, plan, out_dir)

        def done(result):
            path, checks = result
            self.last_output = path
            self.result.setRowCount(0)
            for c in checks:
                i = self.result.rowCount()
                self.result.insertRow(i)
                self.result.setItem(i, 0, ro_item(c.name))
                ok = ro_item("OK" if c.ok else "NG")
                ok.setForeground(QBrush(QColor("#1f7a3d" if c.ok else "#b00020")))
                self.result.setItem(i, 1, ok)
                self.result.setItem(i, 2, ro_item(c.detail))
            n_ng = sum(1 for c in checks if not c.ok)
            self.result_label.setText(f"出力しました：{path}　検証 {'すべて OK' if not n_ng else f'NG {n_ng}件'}")
            self.tabs.setCurrentIndex(2)

        if sync:
            done(work(lambda *_: None))
            return self.last_output
        run_in_thread(self, work, done, "Excel出力")
        return None
