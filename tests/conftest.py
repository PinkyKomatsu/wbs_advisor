"""テスト共通の準備。保存先（%LOCALAPPDATA%\\WBSAdvisor の代わり）は一時フォルダにする。"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = Path(__file__).resolve().parent / "fixtures"
sys.path.insert(0, str(ROOT))


@pytest.fixture()
def home(tmp_path, monkeypatch):
    monkeypatch.setenv("WBSADVISOR_HOME", str(tmp_path / "home"))
    return tmp_path / "home"


@pytest.fixture()
def cfg(home):
    from core import config
    return config.load()


@pytest.fixture()
def db(home):
    from core.db import Database
    d = Database()
    yield d
    d.close()


@pytest.fixture()
def loaded(db, cfg):
    """ひな形・実用WBS様式・過去事例 3 件を登録した状態。"""
    from core import workflow
    workflow.register_master(db, FIXTURES / "master.xlsx", cfg)
    workflow.register_output(db, FIXTURES / "wbs_template.xlsm", cfg)
    ids = {}
    attrs = {
        "case_A": {"update_type": "リプレース", "beds": 400, "vendor": "ベンダーX", "systems": ["電子カルテ", "PACS", "検査システム"], "year": 2024},
        "case_B": {"update_type": "バージョンアップ", "beds": 200, "vendor": "ベンダーY", "systems": ["電子カルテ"], "year": 2023},
        "case_C": {"update_type": "リプレース", "beds": 500, "vendor": "ベンダーX", "systems": ["電子カルテ", "PACS"], "year": 2025},
    }
    for name, a in attrs.items():
        ids[name] = workflow.import_case(db, FIXTURES / f"{name}.xlsx", cfg, attrs=a, name=name)
    return ids


def excel_available() -> bool:
    from excel_io import com_writer
    return sys.platform == "win32" and com_writer.excel_available()


def excel_pids() -> set[int]:
    import subprocess
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
                         capture_output=True, text=True, encoding="cp932", errors="replace").stdout
    pids = set()
    for line in out.splitlines():
        parts = [p.strip('"') for p in line.split('","')]
        if len(parts) > 1 and parts[1].isdigit():
            pids.add(int(parts[1]))
    return pids
