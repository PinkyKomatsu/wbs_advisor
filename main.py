"""エントリポイント。

    WBSAdvisor.exe                 画面を起動
    WBSAdvisor.exe --selftest DIR  画面を自動で操作し、様式登録 → 事例取込 → 照合 → 判定 → コピー → 出力 → 検証 を行う
                                   （結果は DIR\\selftest.json、画面の画像も DIR に保存。データは DIR の中に作る）

データ・設定・ログは %LOCALAPPDATA%\\WBSAdvisor\\ に保存する。ログにセルの内容は書かない。外部通信は一切行わない。
"""
from __future__ import annotations

import argparse
import json
import logging
import logging.handlers
import os
import sys
import time
import traceback
from pathlib import Path


def setup_logging():
    from core import paths
    handler = logging.handlers.RotatingFileHandler(paths.logs_dir() / "app.log", maxBytes=2_000_000, backupCount=5,
                                                   encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.addHandler(handler)
    return paths.logs_dir() / "app.log"


def selftest(win, app, out: Path, fixtures: Path) -> dict:
    from PySide6.QtGui import QGuiApplication
    from PySide6.QtWidgets import QApplication
    result = {"ok": False, "steps": {}}

    def pump(ms=300):
        end = time.time() + ms / 1000
        while time.time() < end:
            QApplication.processEvents()
            time.sleep(0.02)

    def shot(name):
        pump()
        win.grab().save(str(out / name))

    try:
        st = win.state
        win.format_page.master.register(fixtures / "master.xlsx")
        win.format_page.output.register(fixtures / "wbs_template.xlsm")
        win.tabs.setCurrentWidget(win.format_page)
        shot("1_format.png")
        result["steps"]["master_items"] = len(st.db.master_items(st.db.active_format("master")["id"]))
        win.cases_page.import_files([fixtures / f"case_{x}.xlsx" for x in "ABC"], sync=True)
        result["steps"]["cases"] = len(st.db.cases())
        win.tabs.setCurrentWidget(win.cases_page)
        shot("2_cases.png")
        win.matching_page.refresh()
        result["steps"]["pending"] = win.matching_page.pending.rowCount()
        win.tabs.setCurrentWidget(win.matching_page)
        shot("3_matching.png")
        win.matching_page.confirm()
        win.judgment_page.attrs.set_attrs({"update_type": "リプレース", "beds": 450, "systems": ["電子カルテ", "PACS"]})
        win.judgment_page.run(sync=True)
        result["steps"]["judgments"] = len(st.judgments)
        win.tabs.setCurrentWidget(win.judgment_page)
        shot("5_judgment.png")
        ja, en = win.judgment_page.copy("ja"), win.judgment_page.copy("en")
        assert QGuiApplication.clipboard().text() == en
        (out / "copy_ja.tsv").write_text(ja, encoding="utf-8")
        (out / "copy_en.tsv").write_text(en, encoding="utf-8")
        win.tabs.setCurrentWidget(win.export_page)
        plan = win.export_page.make_plan()
        result["steps"]["plan"] = {"engine": plan.engine, "rows": plan.rows_needed, "inserts": plan.insert_rows,
                                   "cells": len(plan.actions), "blocked": plan.blocked}
        shot("6_export.png")
        path = win.export_page.export(confirm=False, out_dir=out, sync=True)
        checks = [(win.export_page.result.item(i, 0).text(), win.export_page.result.item(i, 1).text())
                  for i in range(win.export_page.result.rowCount())]
        result["steps"]["output"] = str(path)
        result["steps"]["checks"] = checks
        shot("6_verify.png")
        win.tabs.setCurrentWidget(win.settings_page)
        shot("7_settings.png")
        ok_design = bool(path) and all(ok == "OK" for _, ok in checks)

        # WBS の種類を切り替え、「大分類・中分類・作業項目」→「管理単位（中２）・（中３）・（小）」で出力する
        win.type_box.setCurrentText("テスト")
        result["steps"]["テスト_empty"] = (st.db.active_format("master") is None and not st.db.cases()
                                           and win.judgment_page.table.rowCount() == 0)
        win.format_page.master.register(fixtures / "master_units.xlsx")
        win.format_page.output.register(fixtures / "template_units.xlsx")
        win.cases_page.import_files([fixtures / "units_A.xlsx", fixtures / "units_B.xlsx"], sync=True)
        win.judgment_page.run(sync=True)
        win.tabs.setCurrentWidget(win.judgment_page)
        shot("8_units_judgment.png")
        # 管理単位（大）は過去事例の最多値が初期値。（中１）は手で選び直す
        result["steps"]["テスト_unit_default"] = win.export_page.unit_combos["A"].currentText()
        win.export_page.select_unit("B", "詳細設計")
        plan2 = win.export_page.make_plan()
        path2 = win.export_page.export(confirm=False, out_dir=out / "テスト", sync=True)
        from excel_io import reader as _reader
        ws = _reader.load(path2)["WBS"]
        result["steps"]["テスト_output"] = [[ws[f"{c}{r}"].value for c in "ABCDE"] for r in (4, 9)]
        rs = win.export_page.result
        result["steps"]["テスト_checks"] = [(rs.item(i, 0).text(), rs.item(i, 1).text(), rs.item(i, 2).text())
                                           for i in range(rs.rowCount())]
        checks2 = [c[1] for c in result["steps"]["テスト_checks"]]
        win.tabs.setCurrentWidget(win.export_page)
        shot("9_units_export.png")
        win.type_box.setCurrentText("設計")
        result["steps"]["設計_cases_after_switch_back"] = len(st.db.cases())
        result["ok"] = (ok_design and result["steps"]["テスト_empty"] and all(c == "OK" for c in checks2)
                        # PACS連携仕様の作成は採用率 50%（要検討）で初期状態は出力しないため、9 行目は次の項目
                        and result["steps"]["テスト_unit_default"] == "医事会計更新"
                        and result["steps"]["テスト_output"] == [
                            ["医事会計更新", "詳細設計", "基本設計", "画面設計", "画面一覧の作成"],
                            ["医事会計更新", "詳細設計", "詳細設計", "マスタ設計", "点数マスタ移行設計"]]
                        and result["steps"]["設計_cases_after_switch_back"] == 3)
    except Exception as e:
        result["error"] = f"{e}\n{traceback.format_exc()}"
    (out / "selftest.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--selftest", metavar="DIR")
    ap.add_argument("--fixtures", metavar="DIR")
    args, _ = ap.parse_known_args(argv)
    if args.selftest:
        os.environ["WBSADVISOR_HOME"] = str(Path(args.selftest) / "home")
        Path(args.selftest).mkdir(parents=True, exist_ok=True)
    log_path = setup_logging()
    log = logging.getLogger("wbsadvisor")
    log.info("起動")

    from PySide6.QtCore import Qt
    from PySide6.QtGui import QFont, QFontDatabase, QGuiApplication
    from PySide6.QtWidgets import QApplication, QMessageBox

    QGuiApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv[:1])
    for family in ("Yu Gothic UI", "Meiryo UI"):
        if family in QFontDatabase.families():
            app.setFont(QFont(family, 9))
            break

    def excepthook(exc_type, exc, tb):
        frames = traceback.extract_tb(tb)
        where = " <- ".join(f"{Path(f.filename).name}:{f.lineno}" for f in reversed(frames[-5:]))
        log.error("予期しないエラー: %s at %s", exc_type.__name__, where)   # メッセージ（セルの内容を含み得る）は書かない
        from app.widgets import error_message
        QMessageBox.critical(None, "エラー", f"予期しないエラーが発生しました。\n{error_message(exc)}\n\n記録：{log_path}")

    sys.excepthook = excepthook
    try:
        from app.main_window import MainWindow
        from app.widgets import AppState
        win = MainWindow(AppState())
    except Exception as e:
        excepthook(type(e), e, e.__traceback__)
        return 1
    win.show()
    if args.selftest:
        fixtures = Path(args.fixtures) if args.fixtures else Path(__file__).resolve().parent / "tests" / "fixtures"
        r = selftest(win, app, Path(args.selftest), fixtures)
        win.close()
        return 0 if r["ok"] else 1
    return app.exec()


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
