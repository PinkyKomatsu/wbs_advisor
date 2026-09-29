# -*- mode: python ; coding: utf-8 -*-
# PyInstaller の設定：単一 exe（--onefile 相当）・ウィンドウアプリ（--windowed 相当）
#   .venv\Scripts\pyinstaller --noconfirm --clean build.spec
#   → dist\WBSAdvisor.exe
# resources（初期設定・定型文テンプレート）は exe に同梱し、実行時は一時展開先（_MEIPASS）から読む。
# データ・設定・ログは %LOCALAPPDATA%\WBSAdvisor\ に保存するため、exe の置き場所は書込不可でもよい。

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=[("resources", "resources")],
    hiddenimports=["win32timezone", "lxml._elementpath"],
    excludes=[
        "tkinter", "PIL", "numpy", "pytest",
        "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets", "PySide6.QtQml", "PySide6.QtQuick",
        "PySide6.Qt3DCore", "PySide6.QtMultimedia", "PySide6.QtCharts", "PySide6.QtDataVisualization",
        "PySide6.QtPdf", "PySide6.QtNetwork", "PySide6.QtOpenGL", "PySide6.QtSql", "PySide6.QtTest",
    ],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="WBSAdvisor",
    console=False,
    disable_windowed_traceback=False,
    upx=False,
    runtime_tmpdir=None,
)
