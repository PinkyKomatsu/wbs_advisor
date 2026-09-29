"""4. 条件ルール管理：属性・演算子・値 → 対象項目を 必要/不要/要検討 に設定する。"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QFormLayout, QHBoxLayout, QHeaderView, QLabel,
                               QLineEdit, QListWidget, QListWidgetItem, QMessageBox, QPushButton, QSpinBox, QSplitter,
                               QTableWidget, QVBoxLayout, QWidget)

from core import workflow
from core.judgment import OPERATORS, RULE_ATTRS
from core.models import DECISIONS, Rule

from ..widgets import AppState, guarded, ro_item


class RulesPage(QWidget):
    def __init__(self, state: AppState, parent=None):
        super().__init__(parent)
        self.state = state
        self.current: Rule | None = None

        self.list = QTableWidget(0, 5)
        self.list.setHorizontalHeaderLabels(["有効", "優先順位", "名称", "条件", "判定・対象"])
        self.list.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.list.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.list.verticalHeader().setVisible(False)
        self.list.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.list.itemSelectionChanged.connect(self._selected)
        lb = QHBoxLayout()
        new = QPushButton("新規ルール")
        new.clicked.connect(lambda: self.edit(Rule(None, "新しいルール")))
        dele = QPushButton("削除")
        dele.clicked.connect(lambda: self.delete())
        lb.addWidget(new)
        lb.addWidget(dele)
        lb.addStretch(1)
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.addWidget(QLabel("ルールは統計判定より優先されます。複数当てはまる場合は優先順位の小さいものを使います。"))
        ll.addWidget(self.list, 1)
        ll.addLayout(lb)

        editor = QWidget()
        form = QFormLayout(editor)
        self.name = QLineEdit()
        self.name_en = QLineEdit()
        self.name_en.setPlaceholderText("判定理由の英文に使います（空欄なら用語辞書で訳します）")
        self.enabled = QCheckBox("有効")
        self.priority = QSpinBox()
        self.priority.setRange(1, 9999)
        self.decision = QComboBox()
        self.decision.addItems(DECISIONS)
        form.addRow("名称", self.name)
        form.addRow("名称（英語）", self.name_en)
        form.addRow("", self.enabled)
        form.addRow("優先順位（小さいほど優先）", self.priority)
        self.conds = QTableWidget(0, 3)
        self.conds.setHorizontalHeaderLabels(["属性", "演算子", "値"])
        self.conds.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.conds.verticalHeader().setVisible(False)
        self.conds.setMaximumHeight(150)
        cb = QHBoxLayout()
        add_c = QPushButton("条件を追加")
        add_c.clicked.connect(lambda: self._add_cond({"attr": "systems", "op": "含む", "value": ""}))
        rm_c = QPushButton("選んだ条件を削除")
        rm_c.clicked.connect(lambda: [self.conds.removeRow(r) for r in sorted({i.row() for i in self.conds.selectedIndexes()}, reverse=True)])
        cb.addWidget(add_c)
        cb.addWidget(rm_c)
        cb.addStretch(1)
        cond_box = QWidget()
        cl = QVBoxLayout(cond_box)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.addWidget(self.conds)
        cl.addLayout(cb)
        form.addRow("条件（すべて満たすとき）", cond_box)
        form.addRow("対象項目を", self.decision)
        self.filter = QLineEdit()
        self.filter.setPlaceholderText("項目名で絞り込み")
        self.filter.textChanged.connect(self._filter)
        self.targets = QListWidget()
        self.targets.setMinimumHeight(220)
        self.keyword = QLineEdit()
        self.keyword.setPlaceholderText("例：PACS（項目名にこの語を含む項目も対象）")
        tb = QWidget()
        tl = QVBoxLayout(tb)
        tl.setContentsMargins(0, 0, 0, 0)
        tl.addWidget(self.filter)
        tl.addWidget(self.targets)
        form.addRow("対象項目", tb)
        form.addRow("項目名に含む語", self.keyword)
        save = QPushButton("ルールを保存")
        save.setStyleSheet("font-weight:bold")
        save.clicked.connect(lambda: self.save())
        form.addRow("", save)

        split = QSplitter(Qt.Horizontal)
        split.addWidget(left)
        split.addWidget(editor)
        split.setSizes([700, 600])
        layout = QVBoxLayout(self)
        layout.addWidget(split)
        state.formatsChanged.connect(self._load_items)
        self._load_items()
        self.refresh()

    def _load_items(self):
        self.targets.clear()
        for i in workflow.master_items(self.state.db):
            it = QListWidgetItem(("　" * (i.level - 1)) + f"{i.wbs_no} {i.name}")
            it.setData(Qt.UserRole, i.key)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Unchecked)
            self.targets.addItem(it)

    def _filter(self, text):
        for n in range(self.targets.count()):
            it = self.targets.item(n)
            it.setHidden(bool(text) and text not in it.text())

    def refresh(self):
        self.rules = self.state.db.rules()
        self.list.setRowCount(0)
        for r in self.rules:
            i = self.list.rowCount()
            self.list.insertRow(i)
            conds = " かつ ".join(f"{RULE_ATTRS.get(c['attr'], c['attr'])} {c['op']} {c.get('value', '')}" for c in r.conditions)
            target = f"{r.decision}（{len(r.targets)}項目" + (f"＋「{r.keyword}」を含む項目" if r.keyword else "") + "）"
            for col, v in enumerate(("○" if r.enabled else "−", r.priority, r.name, conds or "（常に）", target)):
                self.list.setItem(i, col, ro_item(v, r.id))

    def _selected(self):
        rows = {i.row() for i in self.list.selectedIndexes()}
        if rows:
            rid = self.list.item(min(rows), 0).data(Qt.UserRole)
            self.edit(next(r for r in self.rules if r.id == rid))

    def _add_cond(self, c):
        i = self.conds.rowCount()
        self.conds.insertRow(i)
        attr = QComboBox()
        for k, label in RULE_ATTRS.items():
            attr.addItem(label, k)
        attr.setCurrentIndex(max(0, attr.findData(c.get("attr"))))
        op = QComboBox()
        op.addItems(OPERATORS)
        op.setCurrentText(c.get("op", "＝"))
        self.conds.setCellWidget(i, 0, attr)
        self.conds.setCellWidget(i, 1, op)
        value = QLineEdit(str(c.get("value", "")))
        self.conds.setCellWidget(i, 2, value)

    def edit(self, rule: Rule):
        self.current = rule
        self.name.setText(rule.name)
        self.name_en.setText(rule.name_en)
        self.enabled.setChecked(rule.enabled)
        self.priority.setValue(rule.priority)
        self.decision.setCurrentText(rule.decision)
        self.keyword.setText(rule.keyword)
        self.conds.setRowCount(0)
        for c in rule.conditions:
            self._add_cond(c)
        targets = set(rule.targets)
        for n in range(self.targets.count()):
            it = self.targets.item(n)
            it.setCheckState(Qt.Checked if it.data(Qt.UserRole) in targets else Qt.Unchecked)

    def collect(self) -> Rule:
        conds = []
        for i in range(self.conds.rowCount()):
            conds.append({"attr": self.conds.cellWidget(i, 0).currentData(), "op": self.conds.cellWidget(i, 1).currentText(),
                          "value": self.conds.cellWidget(i, 2).text().strip()})
        targets = [self.targets.item(n).data(Qt.UserRole) for n in range(self.targets.count())
                   if self.targets.item(n).checkState() == Qt.Checked]
        r = self.current or Rule(None, "")
        return Rule(r.id, self.name.text().strip() or "無題のルール", self.name_en.text().strip(), self.enabled.isChecked(),
                    self.priority.value(), conds, self.decision.currentText(), targets, self.keyword.text().strip())

    @guarded("ルールを保存できませんでした")
    def save(self):
        rule = self.collect()
        if not rule.targets and not rule.keyword:
            raise ValueError("対象項目を選ぶか、「項目名に含む語」を入力してください。")
        rule.id = self.state.db.save_rule(rule)
        self.current = rule
        self.refresh()
        self.state.judgmentsChanged.emit()

    @guarded("削除できませんでした")
    def delete(self):
        if self.current and self.current.id and \
                QMessageBox.question(self, "削除の確認", f"ルール「{self.current.name}」を削除しますか？") == QMessageBox.Yes:
            self.state.db.delete_rule(self.current.id)
            self.current = None
            self.refresh()
