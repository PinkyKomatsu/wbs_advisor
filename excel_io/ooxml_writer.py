"""OOXML 直接編集方式（Excel がない PC 用のフォールバック）。

テンプレートを zip として扱い、対象シート XML 内の該当セルの値だけを書き換える。
- 文字列は inlineStr、数値は数値、日付は Excel のシリアル値で書く。値を消すときはセル（スタイル）を残して値だけ消す
- セルのスタイル番号（s）は変えない（書式を変更しない）。styles.xml も含め、他のパーツはバイト単位で変更しない
- 行の挿入には対応しない（空き行の不足は書き込み計画の段階で止める）
- 名前空間の接頭辞を保つため、XML は lxml で編集する
"""
from __future__ import annotations

import os
import re
import shutil
import tempfile
import zipfile
from datetime import date, datetime
from pathlib import Path

from lxml import etree

from . import inspect

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
M = "{%s}" % NS
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_REF_RE = re.compile(r"^([A-Z]{1,3})(\d+)$")
_ILLEGAL = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
_PARSER = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def _split(ref: str) -> tuple[int, int]:
    m = _REF_RE.match(ref)
    return _col_index(m.group(1)), int(m.group(2))


def _excel_serial(d) -> float:
    base = datetime(1899, 12, 30)
    if isinstance(d, datetime):
        return (d - base).total_seconds() / 86400
    return float((datetime(d.year, d.month, d.day) - base).days)


def _row(sheet_data, r: int):
    for row in sheet_data.findall(M + "row"):
        n = int(row.get("r"))
        if n == r:
            return row
        if n > r:
            new = etree.Element(M + "row", r=str(r))
            row.addprevious(new)
            return new
    return etree.SubElement(sheet_data, M + "row", r=str(r))


def _cell(row, ref: str):
    col, r = _split(ref)
    for c in row.findall(M + "c"):
        n = _split(c.get("r"))[0]
        if n == col:
            return c
        if n > col:
            new = etree.Element(M + "c", r=ref)
            c.addprevious(new)
            return new
    new = etree.Element(M + "c", r=ref)
    ext = row.find(M + "extLst")
    if ext is not None:
        ext.addprevious(new)
    else:
        row.append(new)
    return new


def set_value(sheet_data, ref: str, value) -> None:
    row = _row(sheet_data, _split(ref)[1])
    row.attrib.pop("spans", None)
    c = _cell(row, ref)
    if c.find(M + "f") is not None:
        raise ValueError(f"{ref} は数式セルのため書き込めません")
    for child in list(c):
        if child.tag != M + "extLst":
            c.remove(child)
    c.attrib.pop("t", None)
    if value is None or value == "":
        return
    if isinstance(value, bool):
        c.set("t", "b")
        etree.SubElement(c, M + "v").text = "1" if value else "0"
    elif isinstance(value, (int, float)):
        etree.SubElement(c, M + "v").text = repr(float(value)) if isinstance(value, float) else str(value)
    elif isinstance(value, (datetime, date)):
        etree.SubElement(c, M + "v").text = repr(_excel_serial(value))
    else:
        c.set("t", "inlineStr")
        is_ = etree.SubElement(c, M + "is")
        t = etree.SubElement(is_, M + "t")
        t.text = _ILLEGAL.sub("", str(value))
        t.set(XML_SPACE, "preserve")
    ext = c.find(M + "extLst")
    if ext is not None:     # extLst はセルの最後の子要素
        c.remove(ext)
        c.append(ext)


def write(template, out, sheet: str, values: dict[str, object]) -> Path:
    """template をコピーした新しいファイル out に、values（{セル: 値}、None は値の消去）を書く。"""
    template, out = Path(template), Path(out)
    if out.exists() and out.resolve() == template.resolve():
        raise ValueError("テンプレートは上書きできません")
    with zipfile.ZipFile(template) as zin:
        part = inspect.sheet_part(zin, sheet)
        root = etree.fromstring(zin.read(part), _PARSER)
        sheet_data = root.find(M + "sheetData")
        for ref in sorted(values, key=lambda r: (_split(r)[1], _split(r)[0])):
            set_value(sheet_data, ref, values[ref])
        new_sheet = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        fd, tmp = tempfile.mkstemp(suffix=out.suffix, dir=out.parent)
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp, "w") as zout:
                for info in zin.infolist():
                    data = new_sheet if info.filename == part else zin.read(info.filename)
                    zout.writestr(info, data, compress_type=info.compress_type)
            shutil.move(tmp, out)
        except BaseException:
            if os.path.exists(tmp):
                os.unlink(tmp)
            raise
    return out
