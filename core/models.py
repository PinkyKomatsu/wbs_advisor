"""データの型。"""
from __future__ import annotations

from dataclasses import dataclass, field

DECISIONS = ("必要", "要検討", "不要")

# 列の意味（様式プロファイルで使う）
MEANINGS = {
    "no": "No（連番）",
    "wbs_no": "WBS番号",
    "phase": "工程/フェーズ",
    "unit_l": "管理単位（大）",
    "unit_m1": "管理単位（中１）",
    "cat1": "大分類",
    "cat2": "中分類",
    "name": "作業項目名",
    "owner": "担当",
    "effort": "予定工数",
    "start": "開始日",
    "end": "終了日",
    "status": "状態",
    "note": "備考",
}


@dataclass
class MasterItem:
    key: str
    seq: int
    row: int
    wbs_no: str
    phase: str
    name: str
    norm: str
    level: int = 1
    parent_key: str | None = None
    values: dict = field(default_factory=dict)   # 列の意味 -> 値


@dataclass
class CaseRow:
    id: int | None
    case_id: int | None
    seq: int
    row: int
    wbs_no: str
    phase: str
    name: str
    norm: str
    values: dict = field(default_factory=dict)
    item_key: str | None = None
    method: str | None = None      # wbs / synonym / exact / fuzzy / manual / none
    score: float | None = None
    status: str | None = None      # auto / pending（確認待ち） / confirmed / extra（ひな形外）
    candidates: list = field(default_factory=list)
    prev_key: str | None = None    # 直前に対応付いたひな形項目（ひな形外の項目の出力位置に使う）


@dataclass
class Case:
    id: int
    name: str
    source: str = ""
    sha256: str = ""
    imported: str = ""
    attrs: dict = field(default_factory=dict)


@dataclass
class Rule:
    id: int | None
    name: str
    name_en: str = ""
    enabled: bool = True
    priority: int = 100
    conditions: list = field(default_factory=list)   # [{"attr", "op", "value"}]（すべて満たすとき）
    decision: str = "必要"
    targets: list = field(default_factory=list)      # 対象項目の item_key
    keyword: str = ""                                # 項目名にこの語を含む項目も対象


@dataclass
class MatchResult:
    row_id: int | None
    item_key: str | None
    method: str
    score: float
    status: str
    candidates: list = field(default_factory=list)
    prev_key: str | None = None


@dataclass
class Judgment:
    key: str
    name: str
    wbs_no: str
    phase: str
    level: int
    is_extra: bool
    total: int
    adopted_cases: list
    not_adopted_cases: list
    rate: float | None
    wrate: float | None
    stat_decision: str
    rule_name: str | None = None
    rule_name_en: str | None = None
    rule_decision: str | None = None
    override: str | None = None
    decision: str = "要検討"
    checked: bool = False
    parent_key: str | None = None
    values: dict = field(default_factory=dict)
    prev_key: str | None = None
    name_en: str = ""
    cat1: str = ""          # 大分類
    cat2: str = ""          # 中分類
    cat1_en: str = ""
    cat2_en: str = ""
    reason_ja: str = ""
    reason_en: str = ""
    untranslated: list = field(default_factory=list)
