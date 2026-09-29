"""5. 判定結果：一覧・手動上書き・日英文の表示・日英別々のコピー。"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor, QGuiApplication
from PySide6.QtWidgets import (QAbstractItemView, QButtonGroup, QCheckBox, QComboBox, QGroupBox, QHBoxLayout,
                               QHeaderView, QInputDialog, QLabel, QPlainTextEdit, QPushButton, QRadioButton,
                               QSplitter, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget)

from core import texts, workflow
from core.models import DECISIONS

from ..widgets import AppState, AttrsEditor, guarded, ro_item, run_in_thread

DECISION_COLORS = {"必要": "#e6f4ea", "要検討": "#fff3cd", "不要": "#f1f1f1"}
COLS = ["出力", "WBS番号", "大分類", "中分類", "作業項目", "作業項目（英語）", "判定", "採用率", "加重", "ルール", "採用した事例", "採用しなかった事例"]
DECISION_COL = COLS.index("判定")


class JudgmentPage(QWidget):
    def __init__(self, state: AppState, parent=None):
        super().__init__(parent)
        self.state = state
        self._filling = False

        proj = QGroupBox("今回の案件の属性（重み付けとルールに使います。未入力でも判定できます）")
        pl = QHBoxLayout(proj)
        self.attrs = AttrsEditor(state.cfg)
        self.attrs.set_attrs(state.db.project_attrs())
        pl.addWidget(self.attrs, 1)
        side = QVBoxLayout()
        self.weighting = QCheckBox("属性が近い事例の重みを上げる")
        self.weighting.setChecked(bool(state.cfg.get("weighting", {}).get("enabled", True)))
        run = QPushButton("判定を実行")
        run.setStyleSheet("font-weight:bold; padding:6px 16px")
        run.clicked.connect(lambda: self.run())
        side.addWidget(self.weighting)
        side.addWidget(run)
        side.addStretch(1)
        pl.addLayout(side)

        self.table = QTableWidget(0, len(COLS))
        self.table.setHorizontalHeaderLabels(COLS)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(COLS.index("作業項目"), QHeaderView.Stretch)
        self.table.itemChanged.connect(self._item_changed)
        self.table.itemSelectionChanged.connect(self.update_texts)
        self.summary = QLabel("")
        self.untranslated = QLabel("")
        self.untranslated.setStyleSheet("color:#b35c00")

        # 日英の文章とコピー
        self.ja = QPlainTextEdit()
        self.en = QPlainTextEdit()
        for t in (self.ja, self.en):
            t.setReadOnly(True)
        self.scope_all = QRadioButton("全項目")
        self.scope_sel = QRadioButton("選択行のみ")
        self.scope_all.setChecked(True)
        self.fmt_tsv = QRadioButton("タブ区切り（Excel に貼付）")
        self.fmt_txt = QRadioButton("改行区切りテキスト")
        (self.fmt_tsv if state.cfg.get("copy", {}).get("format", "tsv") == "tsv" else self.fmt_txt).setChecked(True)
        g1, g2 = QButtonGroup(self), QButtonGroup(self)
        for b in (self.scope_all, self.scope_sel):
            g1.addButton(b)
            b.toggled.connect(lambda _: self.update_texts())
        for b in (self.fmt_tsv, self.fmt_txt):
            g2.addButton(b)
            b.toggled.connect(lambda _: self.update_texts())
        copy_ja = QPushButton("日本語をコピー")
        copy_ja.clicked.connect(lambda: self.copy("ja"))
        copy_en = QPushButton("英語をコピー")
        copy_en.clicked.connect(lambda: self.copy("en"))
        opts = QHBoxLayout()
        for w in (QLabel("範囲："), self.scope_all, self.scope_sel, QLabel("　形式："), self.fmt_tsv, self.fmt_txt):
            opts.addWidget(w)
        opts.addStretch(1)
        ja_box, en_box = QWidget(), QWidget()
        for box, label, text, btn in ((ja_box, "日本語", self.ja, copy_ja), (en_box, "英語（English）", self.en, copy_en)):
            l = QVBoxLayout(box)
            l.setContentsMargins(0, 0, 0, 0)
            head = QHBoxLayout()
            head.addWidget(QLabel(label))
            head.addStretch(1)
            head.addWidget(btn)
            l.addLayout(head)
            l.addWidget(text)
        texts_split = QSplitter(Qt.Horizontal)
        texts_split.addWidget(ja_box)
        texts_split.addWidget(en_box)
        bottom = QWidget()
        bl = QVBoxLayout(bottom)
        bl.setContentsMargins(0, 0, 0, 0)
        bl.addLayout(opts)
        bl.addWidget(texts_split, 1)

        vsplit = QSplitter(Qt.Vertical)
        mid = QWidget()
        ml = QVBoxLayout(mid)
        ml.setContentsMargins(0, 0, 0, 0)
        row = QHBoxLayout()
        row.addWidget(self.summary, 1)
        row.addWidget(self.untranslated)
        ml.addLayout(row)
        ml.addWidget(self.table, 1)
        vsplit.addWidget(mid)
        vsplit.addWidget(bottom)
        vsplit.setSizes([520, 260])

        layout = QVBoxLayout(self)
        layout.addWidget(proj)
        layout.addWidget(vsplit, 1)
        state.judgmentsChanged.connect(self.show_judgments)

    # ---- 判定 ----
    def run(self, sync: bool = False):
        st = self.state
        st.db.save_project_attrs(self.attrs.attrs())
        st.cfg.setdefault("weighting", {})["enabled"] = self.weighting.isChecked()

        def work(progress):
            progress(10, "判定しています")
            return workflow.run_judgment(st.db, st.cfg, st.templates)

        def done(result):
            st.judgments, st.untranslated = result
            st.judgmentsChanged.emit()

        if sync:
            done(work(lambda *_: None))
        else:
            run_in_thread(self, work, done, "判定")

    def show_judgments(self):
        js = self.state.judgments
        self._filling = True
        self.table.setRowCount(0)
        has_cats = any(j.cat1 or j.cat2 for j in js)
        for c in (COLS.index("大分類"), COLS.index("中分類")):
            self.table.setColumnHidden(c, not has_cats)      # 大分類・中分類のない WBS では隠す
        for c in (COLS.index("WBS番号"),):
            self.table.setColumnHidden(c, bool(js) and not any(j.wbs_no for j in js))
        for j in js:
            i = self.table.rowCount()
            self.table.insertRow(i)
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable | Qt.ItemIsSelectable)
            chk.setCheckState(Qt.Checked if j.checked else Qt.Unchecked)
            chk.setData(Qt.UserRole, j.key)
            self.table.setItem(i, 0, chk)
            name = ("　" * max(0, j.level - 1)) + j.name + ("（ひな形外）" if j.is_extra else "")
            rate = f"{texts.pct(j.rate)}%（{len(j.adopted_cases)}/{j.total}）" if j.total else "-"
            wrate = f"{texts.pct(j.wrate)}%" if j.wrate is not None else "-"
            rule = f"{j.rule_name}→{j.rule_decision}" if j.rule_name else ""
            vals = [j.wbs_no, j.cat1, j.cat2, name, j.name_en, None, rate, wrate, rule, "、".join(j.adopted_cases),
                    "、".join(j.not_adopted_cases)]
            for col, v in enumerate(vals, 1):
                if col == DECISION_COL:
                    continue
                item = ro_item(v)
                if col == COLS.index("作業項目（英語）") and j.untranslated:
                    item.setForeground(QBrush(QColor("#b35c00")))
                self.table.setItem(i, col, item)
            combo = QComboBox()
            combo.addItems(DECISIONS)
            combo.setCurrentText(j.decision)
            combo.setStyleSheet(f"background:{DECISION_COLORS.get(j.decision, '#fff')}")
            combo.currentTextChanged.connect(lambda v, key=j.key, prev=j.decision: self.override(key, v, prev))
            self.table.setCellWidget(i, DECISION_COL, combo)
            tip = j.reason_ja + "\n" + j.reason_en
            for col in range(len(COLS)):
                if self.table.item(i, col):
                    self.table.item(i, col).setToolTip(tip)
        self._filling = False
        n = {d: sum(1 for j in js if j.decision == d) for d in DECISIONS}
        self.summary.setText(f"必要 {n['必要']}／要検討 {n['要検討']}／不要 {n['不要']}　出力対象 {sum(1 for j in js if j.checked)}項目"
                             f"（ひな形外 {sum(1 for j in js if j.is_extra)}項目を含む候補）")
        un = self.state.untranslated
        self.untranslated.setText(f"未翻訳の語が {len(un)}件あります（英文に【未翻訳: …】と表示。設定の用語辞書で登録できます）" if un else "")
        self.update_texts()

    def _item_changed(self, item):
        if self._filling or item.column() != 0:
            return
        key = item.data(Qt.UserRole)
        for j in self.state.judgments:
            if j.key == key:
                j.checked = item.checkState() == Qt.Checked
        self.update_texts()

    @guarded("判定を変更できませんでした")
    def override(self, key: str, decision: str, previous: str):
        if self._filling or decision == previous:
            return
        note, ok = QInputDialog.getText(self, "手動上書き", f"判定を「{previous}」から「{decision}」に変更します。理由（任意）：")
        if not ok:
            self.show_judgments()
            return
        self.state.db.add_override(key, decision, previous, note)
        self.run(sync=True)

    # ---- 文章・コピー ----
    def selected(self):
        if self.scope_sel.isChecked():
            rows = sorted({i.row() for i in self.table.selectedIndexes()})
            keys = [self.table.item(r, 0).data(Qt.UserRole) for r in rows]
            by = {j.key: j for j in self.state.judgments}
            return [by[k] for k in keys if k in by]
        return list(self.state.judgments)

    def text_for(self, lang: str) -> str:
        fmt = "tsv" if self.fmt_tsv.isChecked() else "text"
        return texts.copy_text(self.selected(), lang, fmt, self.state.templates)

    def update_texts(self):
        if self.scope_sel.isChecked() or len(self.state.judgments) < 2000:
            self.ja.setPlainText(self.text_for("ja"))
            self.en.setPlainText(self.text_for("en"))

    def copy(self, lang: str) -> str:
        text = self.text_for(lang)
        QGuiApplication.clipboard().setText(text)
        win = self.window()
        if hasattr(win, "statusBar"):
            win.statusBar().showMessage(("日本語" if lang == "ja" else "英語") + f"をコピーしました（{len(self.selected())}項目）", 5000)
        return text
