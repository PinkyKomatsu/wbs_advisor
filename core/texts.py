"""日本語・英語の文章生成（オフライン）。

- 項目名の英訳：用語辞書（日本語 → 英語）。項目名そのものがあればそれを使い、なければ辞書の語で最長一致の置き換えをする。
  辞書にない部分は【未翻訳: ○○】と表示し、件数を数える。
- 判定理由：日英の定型文テンプレートに変数を差し込む。
- コピー用：タブ区切り（Excel に貼付可能）または改行区切りテキスト。
"""
from __future__ import annotations

import re
from decimal import ROUND_HALF_UP, Decimal

from .models import Judgment

_JA_RE = re.compile(r"[぀-ヿ㐀-鿿ｦ-ﾟ々〆]")


def pct(rate: float | None) -> int:
    """採用率を四捨五入した整数の % にする（87.5 → 88）。"""
    if rate is None:
        return 0
    return int(Decimal(str(rate * 100)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


class Translator:
    def __init__(self, glossary: dict[str, str]):
        self.exact = {ja: en for ja, en in glossary.items() if en}
        self.terms = sorted(self.exact.items(), key=lambda t: -len(t[0]))

    def translate(self, text: str) -> tuple[str, list[str]]:
        """(英語, 未翻訳の語のリスト)。"""
        text = (text or "").strip()
        if not text:
            return "", []
        if text in self.exact:
            return self.exact[text], []
        pieces: list[str] = []
        untranslated: list[str] = []
        i, buf = 0, ""

        def flush():
            nonlocal buf
            if buf:
                if _JA_RE.search(buf):
                    chunk = buf.strip(" 　の・")
                    if chunk:
                        untranslated.append(chunk)
                        pieces.append(f"【未翻訳: {chunk}】")
                elif buf.strip():
                    pieces.append(buf.strip())
                buf = ""

        while i < len(text):
            for ja, en in self.terms:
                if text.startswith(ja, i):
                    flush()
                    pieces.append(en)
                    i += len(ja)
                    break
            else:
                if text[i] == "の" and pieces and not buf:   # 「A の B」の「の」は落とす
                    i += 1
                    continue
                buf += text[i]
                i += 1
        flush()
        return " ".join(pieces), untranslated


def _fmt(template: str, **kw) -> str:
    try:
        return template.format(**kw)
    except (KeyError, IndexError, ValueError):
        return template


def build_texts(judgments: list[Judgment], glossary: dict[str, str], templates: dict) -> dict:
    """各項目に name_en / reason_ja / reason_en / untranslated を設定する。未翻訳の語の一覧を返す。"""
    tr = Translator(glossary)
    dec_en = templates.get("decisions", {})
    all_untranslated: dict[str, int] = {}
    for j in judgments:
        j.name_en, un = tr.translate(j.name)
        ja, en = [], []
        if j.is_extra:
            ja.append(templates["extra"]["ja"])
            en.append(templates["extra"]["en"])
        if j.total == 0:
            ja.append(templates["no_cases"]["ja"])
            en.append(templates["no_cases"]["en"])
        else:
            adopted = len(j.adopted_cases)
            key = "stat" if adopted else "stat_none"
            vals = dict(total=j.total, adopted=adopted, rate=pct(j.rate))
            ja.append(_fmt(templates[key]["ja"], **vals))
            en.append(_fmt(templates[key]["en"], **vals))
            if j.wrate is not None:
                ja.append(_fmt(templates["weighted"]["ja"], wrate=pct(j.wrate)))
                en.append(_fmt(templates["weighted"]["en"], wrate=pct(j.wrate)))
        if j.rule_name:
            rule_en = j.rule_name_en
            if not rule_en:
                rule_en, un2 = tr.translate(j.rule_name)
                un = un + un2
            ja.append(_fmt(templates["rule"]["ja"], rule=j.rule_name))
            en.append(_fmt(templates["rule"]["en"], rule=rule_en))
        if j.override:
            ja.append(_fmt(templates["manual"]["ja"], decision=j.override))
            en.append(_fmt(templates["manual"]["en"], decision=dec_en.get(j.override, j.override)))
        j.reason_ja = "".join(ja)
        j.reason_en = " ".join(en)
        j.untranslated = un
        for u in un:
            all_untranslated[u] = all_untranslated.get(u, 0) + 1
    return all_untranslated


# ---------------------------------------------------------------------------
# コピー用
# ---------------------------------------------------------------------------

HEAD_JA = ["WBS番号", "作業項目", "判定", "採用率", "判定理由"]
HEAD_EN = ["WBS No.", "Work item", "Decision", "Adoption rate", "Reason"]


def copy_text(judgments: list[Judgment], lang: str, fmt: str, templates: dict, header: bool = True) -> str:
    """lang: "ja" / "en"、fmt: "tsv"（タブ区切り） / "text"（改行区切り）。"""
    dec_en = templates.get("decisions", {})
    rows = []
    for j in judgments:
        rate = f"{pct(j.rate)}%" if j.total else "-"
        if lang == "en":
            rows.append([j.wbs_no, j.name_en or j.name, dec_en.get(j.decision, j.decision), rate, j.reason_en])
        else:
            rows.append([j.wbs_no, j.name, j.decision, rate, j.reason_ja])
    if fmt == "tsv":
        lines = (["\t".join(HEAD_EN if lang == "en" else HEAD_JA)] if header else [])
        lines += ["\t".join(str(c).replace("\t", " ").replace("\n", " ") for c in r) for r in rows]
        return "\n".join(lines)
    out = []
    for no, name, dec, rate, reason in rows:
        head = f"{no} {name}".strip()
        out.append(f"{head}: {dec} ({rate}) {reason}" if lang == "en" else f"{head}：{dec}（{rate}）{reason}")
    return "\n".join(out)
