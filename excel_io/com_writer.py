"""Excel COM 方式（Excel がインストールされている場合。推奨・既定）。

- テンプレートをファイルコピーしてから、そのコピーを Excel で開いて書き込む（テンプレートは開かない）
- 不可視の新しい Excel を起動し（利用者が開いている Excel には触れない）、
  AutomationSecurity = msoAutomationSecurityForceDisable、EnableEvents = False、DisplayAlerts = False にして
  マクロやイベントが実行されないようにする
- 行の挿入は、直上行の書式・入力規則・数式をコピーする（コピーされた定数は消す）
- 元と同じ形式（.xlsm なら xlOpenXMLWorkbookMacroEnabled）で保存する
- 処理後は必ず Excel を終了させ、プロセスが残っていれば強制終了する（例外時も含む）
"""
from __future__ import annotations

import logging
import re
import shutil
from datetime import date, datetime
from pathlib import Path

log = logging.getLogger("wbsadvisor.com")

MSO_FORCE_DISABLE = 3
XL_SHIFT_DOWN = -4121
XL_FORMAT_FROM_LEFT_OR_ABOVE = 0
XL_CELL_TYPE_CONSTANTS = 2
FILE_FORMATS = {".xlsx": 51, ".xlsm": 52}
_LOOKS_TYPED = re.compile(r"^\s*([+-]?\d[\d,]*(\.\d+)?%?|\d{1,4}[/\-]\d{1,2}([/\-]\d{1,4})?|=.*|TRUE|FALSE)\s*$", re.I)


def excel_available() -> bool:
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"))
        return True
    except OSError:
        return False


class ExcelSession:
    """不可視の Excel を起動し、終了時にプロセスが残らないようにする。"""

    def __enter__(self):
        import pythoncom
        import win32com.client
        import win32process
        pythoncom.CoInitialize()
        self.xl = win32com.client.DispatchEx("Excel.Application")
        try:
            _, self.pid = win32process.GetWindowThreadProcessId(self.xl.Hwnd)
        except Exception:
            self.pid = None
        self.xl.Visible = False
        self.xl.DisplayAlerts = False
        self.xl.ScreenUpdating = False
        self.xl.EnableEvents = False
        self.xl.AskToUpdateLinks = False
        self.xl.AutomationSecurity = MSO_FORCE_DISABLE
        # Application への参照はこのオブジェクトだけが持つ（呼び出し側が持つと Quit しても Excel が終了しない）
        return self

    def __exit__(self, *exc):
        import gc

        import pythoncom
        try:
            for wb in list(self.xl.Workbooks):
                wb.Close(SaveChanges=False)
            wb = None
        except Exception:
            pass
        gc.collect()          # COM の参照を残さない（残っていると Excel が終了しない）
        try:
            self.xl.Quit()
        except Exception:
            pass
        self.xl = None
        gc.collect()
        _wait_or_kill(self.pid)
        pythoncom.CoUninitialize()
        return False


def _wait_or_kill(pid, timeout_ms: int = 10000) -> None:
    if not pid:
        return
    import win32api
    import win32con
    import win32event
    try:
        h = win32api.OpenProcess(win32con.SYNCHRONIZE | win32con.PROCESS_TERMINATE, False, pid)
    except Exception:
        return   # すでに終了している
    try:
        if win32event.WaitForSingleObject(h, timeout_ms) != win32event.WAIT_OBJECT_0:
            log.warning("Excel プロセスが終了しないため強制終了しました")
            win32api.TerminateProcess(h, 1)
    finally:
        win32api.CloseHandle(h)


def _com_value(value, number_format: str):
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    if isinstance(value, str) and number_format != "@" and _LOOKS_TYPED.match(value):
        return "'" + value     # 「1.1」「10/3」などを数値・日付に変換させない
    return value


VBA_PARTS = re.compile(r"^xl/vbaProject(Signature\w*)?\.bin$")


def restore_vba(template: Path, out: Path) -> bool:
    """Excel が保存し直した vbaProject.bin（と署名）を、テンプレートの原本に差し戻す。
    VBA のコード・シート構成は変えていないので原本がそのまま有効（Excel はキャッシュ部分を書き直すだけ）。"""
    import os
    import tempfile
    import zipfile
    with zipfile.ZipFile(template) as zt:
        originals = {n: zt.read(n) for n in zt.namelist() if VBA_PARTS.match(n)}
    if not originals:
        return False
    fd, tmp = tempfile.mkstemp(suffix=out.suffix, dir=out.parent)
    os.close(fd)
    with zipfile.ZipFile(out) as zin, zipfile.ZipFile(tmp, "w") as zout:
        names = set(zin.namelist())
        if set(originals) - names:
            os.unlink(tmp)
            return False     # Excel がパーツ構成を変えた場合は差し戻さない（検証で報告される）
        for info in zin.infolist():
            data = originals.get(info.filename)
            zout.writestr(info, data if data is not None else zin.read(info.filename), compress_type=info.compress_type)
    shutil.move(tmp, out)
    return True


def write(template, out, plan) -> Path:
    template, out = Path(template), Path(out)
    if out.exists() and out.resolve() == template.resolve():
        raise ValueError("テンプレートは上書きできません")
    shutil.copy2(template, out)
    with ExcelSession() as session:
        _write_in_excel(session.xl, out, plan)
    restore_vba(template, out)
    return out


def _write_in_excel(xl, out: Path, plan) -> None:
    """Excel の中での書き込み。COM の参照はこの関数を抜けると解放される（Excel が終了できるように）。"""
    wb = xl.Workbooks.Open(str(out), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True, AddToMru=False)
    try:
        ws = wb.Worksheets(plan.sheet)
        if plan.insert_rows:
            if ws.ProtectContents:
                raise RuntimeError("シートが保護されているため行を挿入できません")
            at, n = plan.insert_at, plan.insert_rows
            ws.Rows(f"{at}:{at + n - 1}").Insert(Shift=XL_SHIFT_DOWN, CopyOrigin=XL_FORMAT_FROM_LEFT_OR_ABOVE)
            src = plan.source_row if plan.source_row < at else plan.source_row + n
            ws.Rows(src).Copy(ws.Rows(f"{at}:{at + n - 1}"))
            try:
                ws.Rows(f"{at}:{at + n - 1}").SpecialCells(XL_CELL_TYPE_CONSTANTS).ClearContents()
            except Exception:
                pass   # 定数がなければ SpecialCells はエラーになる
            height = ws.Rows(src).RowHeight
            ws.Rows(f"{at}:{at + n - 1}").RowHeight = height
        for w in plan.actions:
            cell = ws.Range(w.ref)
            if cell.HasFormula:
                continue   # 計画で除外済みだが念のため
            if w.status == "clear":
                cell.ClearContents()
            else:
                cell.Value = _com_value(w.value, str(cell.NumberFormat))
            cell = None
        wb.Save()
    finally:
        wb.Close(SaveChanges=False)
        ws = wb = None
