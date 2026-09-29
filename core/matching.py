"""実用WBSの各行を、ひな形の項目に対応付ける。

優先順位
  1. WBS 番号の一致（名称がかけ離れている場合は採用しない）
  2. 同義語辞書（確認画面で選んだ結果）
  3. 作業項目名の完全一致（全角半角・空白・記号を正規化したうえで比較）
  4. rapidfuzz による類似度：auto 以上は自動確定、review〜auto は確認待ち、review 未満はひな形外
"""
from __future__ import annotations

from rapidfuzz import fuzz, process

from .models import CaseRow, MasterItem, MatchResult
from .normalize import norm, norm_wbs


class Matcher:
    def __init__(self, items: list[MasterItem], synonyms: dict[str, str] | None = None, cfg: dict | None = None):
        cfg = cfg or {}
        self.auto = float(cfg.get("auto", 90))
        self.review = float(cfg.get("review", 70))
        self.wbs_name_min = float(cfg.get("wbs_name_min", 60))
        self.phase_bonus = float(cfg.get("phase_bonus", 3))
        self.items = items
        self.by_key = {i.key: i for i in items}
        self.by_wbs: dict[str, MasterItem] = {}
        for i in items:
            w = norm_wbs(i.wbs_no)
            if w and w not in self.by_wbs:
                self.by_wbs[w] = i
        self.by_norm: dict[str, list[MasterItem]] = {}
        for i in items:
            self.by_norm.setdefault(i.norm, []).append(i)
        self.synonyms = synonyms or {}
        self._choices = {i.key: i.norm for i in items}

    def _score(self, a: str, b: str) -> float:
        return float(fuzz.ratio(a, b))

    def match_row(self, row: CaseRow) -> MatchResult:
        n = row.norm or norm(row.name)
        if not n:
            return MatchResult(row.id, None, "none", 0.0, "extra")
        # 1. WBS 番号
        w = norm_wbs(row.wbs_no)
        item = self.by_wbs.get(w) if w else None
        if item is not None:
            s = self._score(n, item.norm)
            if s >= self.wbs_name_min:
                return MatchResult(row.id, item.key, "wbs", s, "auto")
        # 2. 同義語辞書
        key = self.synonyms.get(n)
        if key and key in self.by_key:
            return MatchResult(row.id, key, "synonym", 100.0, "auto")
        # 3. 名称の完全一致
        same = self.by_norm.get(n, [])
        if len(same) == 1:
            return MatchResult(row.id, same[0].key, "exact", 100.0, "auto")
        if len(same) > 1:
            by_phase = [i for i in same if norm(i.phase) == norm(row.phase)]
            if len(by_phase) == 1:
                return MatchResult(row.id, by_phase[0].key, "exact", 100.0, "auto")
            return MatchResult(row.id, None, "exact", 100.0, "pending",
                               [[i.key, 100.0] for i in same][:5])
        # 4. 類似度
        hits = process.extract(n, self._choices, scorer=fuzz.ratio, limit=5)
        cands = []
        for _choice, score, key in hits:
            if norm(self.by_key[key].phase) == norm(row.phase):
                score = min(100.0, score + self.phase_bonus)
            cands.append([key, round(float(score), 1)])
        cands.sort(key=lambda c: -c[1])
        if cands and cands[0][1] >= self.auto:
            return MatchResult(row.id, cands[0][0], "fuzzy", cands[0][1], "auto", cands)
        if cands and cands[0][1] >= self.review:
            return MatchResult(row.id, None, "fuzzy", cands[0][1], "pending", cands)
        return MatchResult(row.id, None, "none", cands[0][1] if cands else 0.0, "extra", cands)

    def match_rows(self, rows: list[CaseRow]) -> list[MatchResult]:
        """1 件の事例の全行を照合する。ひな形外の行には、直前に対応付いた項目を prev_key として記録する。"""
        out, prev = [], None
        for r in sorted(rows, key=lambda r: r.seq):
            m = self.match_row(r)
            m.prev_key = prev
            if m.item_key:
                prev = m.item_key
            out.append(m)
        return out
