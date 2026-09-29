"""必要項目の判定。

- 採用：ひな形の項目が実用WBSに転記されていれば、その事例では「採用」。子の項目が採用されていれば親も採用とみなす。
- 採用率 = 採用した事例数 ÷ 対象事例数。案件属性が近い事例ほど重みを上げた加重採用率も出す（ON/OFF 可）。
- 判定：採用率が required 以上＝必要、unneeded 未満＝不要、その間＝要検討。
- 優先順位：手動上書き ＞ ユーザー独自条件（ルール） ＞ 統計判定。
- ひな形外の追加項目は、名称ごとに事例をまたいで集計し、候補として出す。
"""
from __future__ import annotations

from collections import Counter, defaultdict

from .models import Case, CaseRow, Judgment, MasterItem, Rule
from .normalize import norm

REQUIRED, REVIEW, UNNEEDED = "必要", "要検討", "不要"


# ---------------------------------------------------------------------------
# 案件属性の類似度と重み
# ---------------------------------------------------------------------------

def bed_band(beds, bands: list[int]) -> int | None:
    try:
        b = int(beds)
    except (TypeError, ValueError):
        return None
    return sum(1 for x in bands if b >= x)


def _as_set(v) -> set[str]:
    if v is None or v == "":
        return set()
    if isinstance(v, (list, tuple, set)):
        return {norm(x) for x in v if str(x).strip()}
    return {norm(x) for x in str(v).replace("、", ",").split(",") if x.strip()}


def attr_similarity(case_attrs: dict, project: dict, cfg: dict) -> float | None:
    """0〜1。両方に入力がある属性だけで計算する。比べられる属性がなければ None。"""
    w = cfg.get("weighting", {})
    parts: list[tuple[float, float]] = []

    def filled(v):
        return v not in (None, "", [], ())

    if filled(case_attrs.get("update_type")) and filled(project.get("update_type")):
        parts.append((w.get("update_type", 0.3), 1.0 if case_attrs["update_type"] == project["update_type"] else 0.0))
    bands = cfg.get("bed_bands", [100, 200, 400])
    cb, pb = bed_band(case_attrs.get("beds"), bands), bed_band(project.get("beds"), bands)
    if cb is not None and pb is not None:
        parts.append((w.get("bed_band", 0.2), {0: 1.0, 1: 0.5}.get(abs(cb - pb), 0.0)))
    if filled(case_attrs.get("vendor")) and filled(project.get("vendor")):
        parts.append((w.get("vendor", 0.2), 1.0 if norm(case_attrs["vendor"]) == norm(project["vendor"]) else 0.0))
    for attr in ("systems", "tags"):
        a, b = _as_set(case_attrs.get(attr)), _as_set(project.get(attr))
        if a or b:
            if a and b:
                parts.append((w.get(attr, 0.2), len(a & b) / len(a | b)))
    try:
        dy = abs(int(case_attrs.get("year")) - int(project.get("year")))
        parts.append((w.get("year", 0.05), max(0.0, 1 - dy / max(1, w.get("year_span", 5)))))
    except (TypeError, ValueError):
        pass
    total = sum(p[0] for p in parts)
    if not parts or total <= 0:
        return None
    return sum(p[0] * p[1] for p in parts) / total


def case_weights(cases: list[Case], project: dict, cfg: dict) -> dict[int, float]:
    """事例の重み = 1 + coef × 属性の類似度（比べられない事例は 1）。"""
    w = cfg.get("weighting", {})
    coef = float(w.get("coef", 1.0))
    out = {}
    for c in cases:
        s = attr_similarity(c.attrs or {}, project or {}, cfg) if w.get("enabled", True) else None
        out[c.id] = 1.0 + coef * s if s is not None else 1.0
    return out


# ---------------------------------------------------------------------------
# ユーザー独自条件（ルール）
# ---------------------------------------------------------------------------

OPERATORS = ["＝", "≠", "含む", "含まない", "以上", "以下", "入力あり", "未入力"]
RULE_ATTRS = {"update_type": "更新種別", "beds": "病床規模", "vendor": "ベンダー名", "systems": "連携システム",
              "year": "実施年度", "tags": "自由タグ"}


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def condition_holds(cond: dict, project: dict) -> bool:
    attr, op, value = cond.get("attr"), cond.get("op"), cond.get("value")
    actual = project.get(attr)
    if op == "入力あり":
        return actual not in (None, "", [], ())
    if op == "未入力":
        return actual in (None, "", [], ())
    if actual in (None, "", [], ()):
        return False
    if op in ("含む", "含まない"):
        hit = bool(_as_set(value) & _as_set(actual)) if isinstance(actual, (list, tuple, set)) \
            else norm(value) in norm(actual)
        return hit if op == "含む" else not hit
    if op in ("以上", "以下"):
        a, b = _num(actual), _num(value)
        if a is None or b is None:
            return False
        return a >= b if op == "以上" else a <= b
    eq = _as_set(actual) == _as_set(value) if isinstance(actual, (list, tuple, set)) else norm(actual) == norm(value)
    return eq if op == "＝" else not eq


def rule_applies(rule: Rule, project: dict) -> bool:
    return rule.enabled and all(condition_holds(c, project) for c in rule.conditions)


def rule_targets(rule: Rule, key: str, name: str) -> bool:
    return key in rule.targets or (bool(rule.keyword) and norm(rule.keyword) in norm(name))


# ---------------------------------------------------------------------------
# 判定
# ---------------------------------------------------------------------------

def stat_decision(rate: float | None, cfg: dict) -> str:
    if rate is None:
        return REVIEW
    t = cfg.get("thresholds", {})
    if rate >= float(t.get("required", 0.7)) - 1e-9:
        return REQUIRED
    if rate < float(t.get("unneeded", 0.3)) - 1e-9:
        return UNNEEDED
    return REVIEW


def _ancestors(items: list[MasterItem]) -> dict[str, list[str]]:
    by_key = {i.key: i for i in items}
    out = {}
    for i in items:
        chain, p = [], i.parent_key
        while p and p in by_key and p not in chain:
            chain.append(p)
            p = by_key[p].parent_key
        out[i.key] = chain
    return out


def judge(items: list[MasterItem], cases: list[Case], rows: list[CaseRow], rules: list[Rule],
          project: dict, overrides: dict[str, str], cfg: dict) -> list[Judgment]:
    """ひな形の全項目（ひな形順）と、ひな形外の追加項目を判定する。"""
    project = project or {}
    weights = case_weights(cases, project, cfg)
    use_weight = cfg.get("weighting", {}).get("enabled", True) and any(abs(w - 1.0) > 1e-9 for w in weights.values())
    names = {c.id: c.name for c in cases}
    ancestors = _ancestors(items)

    adopted: dict[int, set[str]] = defaultdict(set)
    extras: dict[str, dict] = {}
    for r in rows:
        if r.case_id not in names:
            continue
        if r.item_key and r.status in ("auto", "confirmed"):
            adopted[r.case_id].add(r.item_key)
            adopted[r.case_id].update(ancestors.get(r.item_key, []))
        elif r.status == "extra" and r.name:
            e = extras.setdefault(r.norm or norm(r.name), {"name": r.name, "phase": r.phase, "cases": set(),
                                                           "prev": Counter(), "values": r.values, "wbs_no": r.wbs_no})
            e["cases"].add(r.case_id)
            if r.prev_key:
                e["prev"][r.prev_key] += 1

    active_rules = sorted([r for r in rules if rule_applies(r, project)], key=lambda r: (r.priority, r.id or 0))
    total = len(cases)
    total_w = sum(weights.values())
    out: list[Judgment] = []

    def build(key, name, wbs_no, phase, level, parent, values, adopting: set[int], is_extra, prev=None):
        rate = len(adopting) / total if total else None
        wrate = sum(weights[c] for c in adopting) / total_w if total and use_weight else None
        sd = stat_decision(wrate if wrate is not None else rate, cfg)
        j = Judgment(key=key, name=name, wbs_no=wbs_no or "", phase=phase or "", level=level, is_extra=is_extra,
                     total=total, adopted_cases=[names[c] for c in sorted(adopting)],
                     not_adopted_cases=[names[c] for c in sorted(names) if c not in adopting],
                     rate=rate, wrate=wrate, stat_decision=sd, parent_key=parent, values=values or {}, prev_key=prev)
        decision = sd
        for rule in active_rules:
            if rule_targets(rule, key, name):
                j.rule_name, j.rule_name_en, j.rule_decision = rule.name, rule.name_en, rule.decision
                decision = rule.decision
                break
        if key in overrides:
            j.override = overrides[key]
            decision = overrides[key]
        j.decision = decision
        j.checked = decision == REQUIRED
        return j

    for i in items:
        adopting = {c for c in names if i.key in adopted[c]}
        out.append(build(i.key, i.name, i.wbs_no, i.phase, i.level, i.parent_key, i.values, adopting, False))
    for n, e in extras.items():
        prev = e["prev"].most_common(1)[0][0] if e["prev"] else None
        out.append(build("extra:" + n, e["name"], e["wbs_no"], e["phase"], 0, None, e["values"], e["cases"], True, prev))
    return out


def output_order(judgments: list[Judgment], include_parents: bool = True) -> list[Judgment]:
    """出力する項目をひな形の順に並べる。子が出力対象なら親も出す。ひな形外の項目は直前の項目の後ろに入れる。"""
    by_key = {j.key: j for j in judgments}
    chosen = {j.key for j in judgments if j.checked}
    if include_parents:
        for j in judgments:
            if j.checked and not j.is_extra:
                p = j.parent_key
                while p and p in by_key:
                    chosen.add(p)
                    p = by_key[p].parent_key
    base = [j for j in judgments if not j.is_extra and j.key in chosen]
    extras = [j for j in judgments if j.is_extra and j.key in chosen]
    result = list(base)
    for e in extras:
        idx = next((k for k, j in enumerate(result) if j.key == e.prev_key), None)
        if idx is None:
            result.append(e)
        else:
            # 同じ親（直前項目）の子孫の後ろ
            end = idx + 1
            while end < len(result) and result[end].level > by_key.get(e.prev_key, e).level:
                end += 1
            result.insert(end, e)
    return result
