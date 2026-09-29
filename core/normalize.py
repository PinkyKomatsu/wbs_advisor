"""照合用の正規化（全角半角・空白・記号）。"""
from __future__ import annotations

import re
import unicodedata

# 比較で無視する空白・記号（長音「ー」はカタカナ語の一部なので残す）
_SYMBOLS_RE = re.compile(r"[\s　・･\-－‐―_/／\\()（）\[\]［］【】「」『』〈〉<>＜＞、。,.．:：;；!！?？\"'`’“”〜~*＊#＃]+")
_WBS_SEP_RE = re.compile(r"[\-－ー―_\s/／]+")


def nfkc(s) -> str:
    return unicodedata.normalize("NFKC", "" if s is None else str(s))


def norm(s) -> str:
    """作業項目名の正規化：NFKC・小文字化・空白と記号を除く。"""
    return _SYMBOLS_RE.sub("", nfkc(s).lower()).strip()


def norm_wbs(s) -> str:
    """WBS 番号の正規化：「2-1」「２．１」「02.01」→「2.1」。"""
    t = nfkc(s).strip()
    if not t:
        return ""
    if re.fullmatch(r"\d+\.0", t):       # 数値として読んだ「1.0」→「1」
        t = t[:-2]
    t = _WBS_SEP_RE.sub(".", t).strip(".")
    parts = [p.lstrip("0") or "0" if p.isdigit() else p for p in t.split(".") if p]
    return ".".join(parts)


def wbs_level(wbs: str) -> int:
    w = norm_wbs(wbs)
    return len(w.split(".")) if w else 1


def item_key(phase: str, name: str) -> str:
    """ひな形項目の識別子（ひな形を登録し直しても同じ項目なら同じ値）。"""
    return f"{norm(phase)}/{norm(name)}"
