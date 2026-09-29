"""7. 設定：閾値・値の対応表・用語辞書・定型文テンプレート。"""
from __future__ import annotations

import csv
import json

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QCheckBox, QComboBox, QDoubleSpinBox, QFileDialog, QFormLayout, QGroupBox, QHBoxLayout,
                               QHeaderView, QLabel, QLineEdit, QMessageBox, QPushButton, QScrollArea, QSpinBox,
                               QTableWidget, QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget)

from core import config as config_mod, paths
from core.models import MEANINGS

from ..widgets import AppState, guarded


def _pct_spin(value: float) -> QSpinBox:
    s = QSpinBox()
    s.setRange(0, 100)
    s.setSuffix(" %")
    s.setValue(int(round(value * 100)))
    return s


class SettingsPage(QWidget):
    def __init__(self, state: AppState, parent=None):
        super().__init__(parent)
        self.state = state
        tabs = QTabWidget()
        tabs.addTab(self._general(), "判定・照合・出力")
        tabs.addTab(self._value_map(), "値の対応表")
        tabs.addTab(self._glossary(), "用語辞書")
        tabs.addTab(self._templates(), "定型文テンプレート")
        layout = QVBoxLayout(self)
        layout.addWidget(tabs)
        info = QLabel(f"データ・設定・ログの保存先：{paths.home()}")
        info.setStyleSheet("color:#666")
        layout.addWidget(info)

    # ---- 判定・照合・出力 ----
    def _general(self):
        c = self.state.cfg
        w = QWidget()
        form = QFormLayout(w)
        self.required = _pct_spin(c["thresholds"]["required"])
        self.unneeded = _pct_spin(c["thresholds"]["unneeded"])
        form.addRow("「必要」とする採用率（以上）", self.required)
        form.addRow("「不要」とする採用率（未満）", self.unneeded)
        self.auto = QSpinBox()
        self.auto.setRange(0, 100)
        self.auto.setValue(int(c["match"]["auto"]))
        self.review = QSpinBox()
        self.review.setRange(0, 100)
        self.review.setValue(int(c["match"]["review"]))
        form.addRow("類似度：自動確定（以上）", self.auto)
        form.addRow("類似度：確認画面に出す（以上）", self.review)
        wt = c["weighting"]
        self.weight_on = QCheckBox("属性が近い事例の重みを上げる")
        self.weight_on.setChecked(bool(wt.get("enabled", True)))
        form.addRow("", self.weight_on)
        self.coef = QDoubleSpinBox()
        self.coef.setRange(0, 10)
        self.coef.setSingleStep(0.1)
        self.coef.setValue(float(wt.get("coef", 1.0)))
        form.addRow("重みの強さ（重み = 1 + 強さ × 類似度）", self.coef)
        self.bands = QLineEdit(", ".join(str(b) for b in c.get("bed_bands", [])))
        form.addRow("病床規模の区分の境目（床）", self.bands)
        self.update_types = QLineEdit("、".join(c["attributes"]["update_types"]))
        self.systems = QLineEdit("、".join(c["attributes"]["systems"]))
        form.addRow("更新種別の選択肢", self.update_types)
        form.addRow("連携システムの選択肢", self.systems)
        self.engine = QComboBox()
        for k, label in (("auto", "自動（Excel があれば COM 方式）"), ("com", "Excel COM 方式"), ("ooxml", "OOXML 直接編集方式")):
            self.engine.addItem(label, k)
        self.engine.setCurrentIndex(max(0, self.engine.findData(c["output"].get("engine", "auto"))))
        form.addRow("書き込み方式", self.engine)
        self.include_parents = QCheckBox("子の項目を出力するときは親（上位階層）の項目も出力する")
        self.include_parents.setChecked(bool(c["output"].get("include_parents", True)))
        self.clear_sample = QCheckBox("テンプレートの明細に残っている値を消してから書き込む")
        self.clear_sample.setChecked(bool(c["output"].get("clear_sample_rows", True)))
        form.addRow("", self.include_parents)
        form.addRow("", self.clear_sample)
        save = QPushButton("設定を保存")
        save.setStyleSheet("font-weight:bold")
        save.clicked.connect(lambda: self.save_general())
        form.addRow("", save)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(w)
        return scroll

    @guarded("設定を保存できませんでした")
    def save_general(self):
        c = self.state.cfg
        if self.unneeded.value() > self.required.value():
            raise ValueError("「不要」の採用率は「必要」の採用率以下にしてください。")
        if self.review.value() > self.auto.value():
            raise ValueError("確認画面に出す類似度は、自動確定の類似度以下にしてください。")
        c["thresholds"] = {"required": self.required.value() / 100, "unneeded": self.unneeded.value() / 100}
        c["match"].update(auto=self.auto.value(), review=self.review.value())
        c["weighting"].update(enabled=self.weight_on.isChecked(), coef=self.coef.value())
        c["bed_bands"] = [int(x) for x in self.bands.text().replace("、", ",").split(",") if x.strip().isdigit()]
        split = lambda s: [x.strip() for x in s.replace(",", "、").split("、") if x.strip()]
        c["attributes"]["update_types"] = split(self.update_types.text())
        c["attributes"]["systems"] = split(self.systems.text())
        c["output"].update(engine=self.engine.currentData(), include_parents=self.include_parents.isChecked(),
                           clear_sample_rows=self.clear_sample.isChecked())
        self.state.save_config()
        QMessageBox.information(self, "設定", "設定を保存しました。選択肢の変更は再起動後の入力欄に反映されます。")

    # ---- 値の対応表 ----
    def _value_map(self):
        w = QWidget()
        l = QVBoxLayout(w)
        l.addWidget(QLabel("入力規則（プルダウン）のあるセルに、許可されていない値を書くときの置き換え表です。"
                           "例：担当「病院」→「病院情報システム部」。対応できない値は書き込まずに警告します。"))
        self.vm = QTableWidget(0, 3)
        self.vm.setHorizontalHeaderLabels(["列の意味（* はすべての列）", "元の値", "書き込む値"])
        self.vm.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        for meaning, table in self.state.cfg.get("value_map", {}).items():
            for src, dst in table.items():
                self._vm_row(meaning, src, dst)
        b = QHBoxLayout()
        add = QPushButton("行を追加")
        add.clicked.connect(lambda: self._vm_row("*", "", ""))
        rm = QPushButton("選んだ行を削除")
        rm.clicked.connect(lambda: [self.vm.removeRow(r) for r in sorted({i.row() for i in self.vm.selectedIndexes()}, reverse=True)])
        save = QPushButton("対応表を保存")
        save.clicked.connect(lambda: self.save_value_map())
        for x in (add, rm):
            b.addWidget(x)
        b.addStretch(1)
        b.addWidget(save)
        l.addWidget(self.vm, 1)
        l.addLayout(b)
        return w

    def _vm_row(self, meaning, src, dst):
        i = self.vm.rowCount()
        self.vm.insertRow(i)
        combo = QComboBox()
        combo.addItem("*（すべての列）", "*")
        for k, label in MEANINGS.items():
            combo.addItem(label, k)
        combo.setCurrentIndex(max(0, combo.findData(meaning)))
        self.vm.setCellWidget(i, 0, combo)
        self.vm.setItem(i, 1, QTableWidgetItem(src))
        self.vm.setItem(i, 2, QTableWidgetItem(dst))

    @guarded("対応表を保存できませんでした")
    def save_value_map(self):
        vm: dict = {"*": {}}
        for i in range(self.vm.rowCount()):
            meaning = self.vm.cellWidget(i, 0).currentData()
            src = (self.vm.item(i, 1) or QTableWidgetItem("")).text().strip()
            dst = (self.vm.item(i, 2) or QTableWidgetItem("")).text().strip()
            if src:
                vm.setdefault(meaning, {})[src] = dst
        self.state.cfg["value_map"] = vm
        self.state.save_config()
        QMessageBox.information(self, "値の対応表", "保存しました。")

    # ---- 用語辞書 ----
    def _glossary(self):
        w = QWidget()
        l = QVBoxLayout(w)
        l.addWidget(QLabel("項目名などの英訳に使います。英語が空欄の語は、英文に【未翻訳: ○○】と表示されます。"))
        self.gfilter = QLineEdit()
        self.gfilter.setPlaceholderText("検索")
        self.gfilter.textChanged.connect(self._gfilter)
        self.untranslated_only = QCheckBox("英語が空欄の語だけ表示")
        self.untranslated_only.toggled.connect(lambda _: self._gfilter(self.gfilter.text()))
        row = QHBoxLayout()
        row.addWidget(self.gfilter, 1)
        row.addWidget(self.untranslated_only)
        l.addLayout(row)
        self.gl = QTableWidget(0, 2)
        self.gl.setHorizontalHeaderLabels(["日本語", "英語"])
        self.gl.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        l.addWidget(self.gl, 1)
        b = QHBoxLayout()
        for text, fn in (("行を追加", self._gl_add), ("CSV をインポート…", self.import_csv), ("CSV をエクスポート…", self.export_csv)):
            x = QPushButton(text)
            x.clicked.connect(lambda _=False, f=fn: f())
            b.addWidget(x)
        b.addStretch(1)
        save = QPushButton("用語辞書を保存")
        save.clicked.connect(lambda: self.save_glossary())
        b.addWidget(save)
        l.addLayout(b)
        self.load_glossary()
        return w

    def load_glossary(self):
        self.gl.setRowCount(0)
        for ja, en in sorted(self.state.db.glossary().items()):
            i = self.gl.rowCount()
            self.gl.insertRow(i)
            self.gl.setItem(i, 0, QTableWidgetItem(ja))
            self.gl.setItem(i, 1, QTableWidgetItem(en))

    def _gfilter(self, text):
        for i in range(self.gl.rowCount()):
            ja, en = self.gl.item(i, 0).text(), (self.gl.item(i, 1).text() if self.gl.item(i, 1) else "")
            hide = (bool(text) and text not in ja and text not in en) or (self.untranslated_only.isChecked() and en)
            self.gl.setRowHidden(i, bool(hide))

    def _gl_add(self):
        i = self.gl.rowCount()
        self.gl.insertRow(i)
        self.gl.setItem(i, 0, QTableWidgetItem(""))
        self.gl.setItem(i, 1, QTableWidgetItem(""))
        self.gl.editItem(self.gl.item(i, 0))

    def _gl_terms(self) -> dict:
        out = {}
        for i in range(self.gl.rowCount()):
            ja = (self.gl.item(i, 0) or QTableWidgetItem("")).text().strip()
            if ja:
                out[ja] = (self.gl.item(i, 1) or QTableWidgetItem("")).text().strip()
        return out

    @guarded("用語辞書を保存できませんでした")
    def save_glossary(self):
        terms = self._gl_terms()
        db = self.state.db
        for ja in set(db.glossary()) - set(terms):
            db.delete_term(ja)
        db.upsert_terms(terms, source="編集", keep_existing_en=False)
        QMessageBox.information(self, "用語辞書", f"保存しました（{len(terms)}語）。判定をやり直すと英文に反映されます。")

    @guarded("インポートできませんでした")
    def import_csv(self, path: str | None = None):
        path = path or QFileDialog.getOpenFileName(self, "用語辞書（CSV：日本語,英語）", "", "CSV (*.csv)")[0]
        if not path:
            return 0
        terms = {}
        for enc in ("utf-8-sig", "cp932"):
            try:
                with open(path, encoding=enc, newline="") as f:
                    for row in csv.reader(f):
                        if len(row) >= 2 and row[0].strip() and row[0].strip() not in ("日本語", "ja"):
                            terms[row[0].strip()] = row[1].strip()
                break
            except UnicodeDecodeError:
                continue
        self.state.db.upsert_terms(terms, source="インポート", keep_existing_en=False)
        self.load_glossary()
        return len(terms)

    @guarded("エクスポートできませんでした")
    def export_csv(self):
        path, _ = QFileDialog.getSaveFileName(self, "用語辞書を保存", "glossary.csv", "CSV (*.csv)")
        if not path:
            return
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["日本語", "英語"])
            for ja, en in sorted(self.state.db.glossary().items()):
                w.writerow([ja, en])

    # ---- 定型文テンプレート ----
    def _templates(self):
        w = QWidget()
        l = QVBoxLayout(w)
        l.addWidget(QLabel("判定理由の定型文です。{total}（事例数）{adopted}（採用数）{rate}（採用率%）{wrate}（加重採用率%）"
                           "{rule}（ルール名）{decision}（判定）が使えます。"))
        self.tt = QTableWidget(0, 3)
        self.tt.setHorizontalHeaderLabels(["種類", "日本語", "英語"])
        self.tt.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tt.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        for key, v in self.state.templates.items():
            if isinstance(v, dict) and "ja" in v:
                i = self.tt.rowCount()
                self.tt.insertRow(i)
                k = QTableWidgetItem(key)
                k.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)   # 種類は変更不可
                self.tt.setItem(i, 0, k)
                self.tt.setItem(i, 1, QTableWidgetItem(v["ja"]))
                self.tt.setItem(i, 2, QTableWidgetItem(v["en"]))
        save = QPushButton("定型文を保存")
        save.clicked.connect(lambda: self.save_templates())
        l.addWidget(self.tt, 1)
        l.addWidget(save)
        return w

    @guarded("定型文を保存できませんでした")
    def save_templates(self):
        t = json.loads(json.dumps(self.state.templates))
        for i in range(self.tt.rowCount()):
            key = self.tt.item(i, 0).text()
            t[key] = {"ja": self.tt.item(i, 1).text(), "en": self.tt.item(i, 2).text()}
        config_mod.save_text_templates(t)
        self.state.templates = config_mod.load_text_templates()
        QMessageBox.information(self, "定型文", "保存しました。判定をやり直すと反映されます。")
