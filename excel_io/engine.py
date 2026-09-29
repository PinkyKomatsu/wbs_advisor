"""書き込みエンジンの選択と出力ファイル名。"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from . import com_writer, ooxml_writer

ENGINE_LABELS = {"com": "Excel COM 方式（Excel で書き込み・行挿入に対応）",
                 "ooxml": "OOXML 直接編集方式（Excel なし・行挿入には未対応）"}


def choose(setting: str = "auto") -> str:
    """auto なら Excel があれば COM、なければ OOXML。"""
    if setting in ("com", "ooxml"):
        if setting == "com" and not com_writer.excel_available():
            return "ooxml"
        return setting
    return "com" if com_writer.excel_available() else "ooxml"


def output_name(template: Path, now: datetime | None = None) -> str:
    now = now or datetime.now()
    return f"{template.stem}_{now:%Y%m%d_%H%M%S}{template.suffix}"


def execute(template, out_dir, plan, now: datetime | None = None) -> Path:
    template = Path(template)
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    out = Path(out_dir) / output_name(template, now)
    if out.resolve() == template.resolve():
        raise ValueError("テンプレートは上書きできません")
    if plan.blocked:
        raise RuntimeError(plan.blocked)
    if plan.engine == "com":
        return com_writer.write(template, out, plan)
    return ooxml_writer.write(template, out, plan.sheet, plan.cell_values())
