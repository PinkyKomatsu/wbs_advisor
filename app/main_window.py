"""メインウィンドウ：7 つの画面をタブで切り替える。重い処理の進捗はステータスバーの進捗バーに出す。"""
from __future__ import annotations

from PySide6.QtWidgets import QLabel, QMainWindow, QProgressBar, QTabWidget

from core import paths
from excel_io import com_writer

from .pages.cases_page import CasesPage
from .pages.export_page import ExportPage
from .pages.format_page import FormatPage
from .pages.judgment_page import JudgmentPage
from .pages.matching_page import MatchingPage
from .pages.rules_page import RulesPage
from .pages.settings_page import SettingsPage
from .widgets import AppState

APP_TITLE = "WBS 項目判定アシスタント（医事会計システム更新）"


class MainWindow(QMainWindow):
    def __init__(self, state: AppState | None = None):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1480, 900)
        self.state = state or AppState()
        self.tabs = QTabWidget()
        self.format_page = FormatPage(self.state)
        self.cases_page = CasesPage(self.state)
        self.matching_page = MatchingPage(self.state)
        self.rules_page = RulesPage(self.state)
        self.judgment_page = JudgmentPage(self.state)
        self.export_page = ExportPage(self.state)
        self.settings_page = SettingsPage(self.state)
        for page, title in ((self.format_page, "1. 様式登録"), (self.cases_page, "2. 過去事例管理"),
                            (self.matching_page, "3. 照合確認"), (self.rules_page, "4. 条件ルール管理"),
                            (self.judgment_page, "5. 判定結果"), (self.export_page, "6. Excel出力"),
                            (self.settings_page, "7. 設定")):
            self.tabs.addTab(page, title)
        self.setCentralWidget(self.tabs)
        self.progress = QProgressBar()
        self.progress.setMaximumWidth(220)
        self.progress.hide()
        self.busy_label = QLabel("")
        self.statusBar().addPermanentWidget(self.busy_label)
        self.statusBar().addPermanentWidget(self.progress)
        excel = "Excel あり（COM 方式）" if com_writer.excel_available() else "Excel なし（OOXML 方式）"
        self.statusBar().showMessage(f"保存先：{paths.home()}　／　{excel}　／　オフライン動作（外部通信なし）")

    def begin_busy(self, title: str):
        self.busy_label.setText(title + "中…")
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.show()

    def set_progress(self, pct: int, text: str = ""):
        self.progress.setValue(max(0, min(100, pct)))
        if text:
            self.busy_label.setText(text)

    def end_busy(self):
        self.progress.hide()
        self.busy_label.setText("")
