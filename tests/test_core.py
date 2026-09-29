"""照合・採用率・閾値判定・ユーザールール・日英テンプレート・未翻訳語のテスト。"""
from __future__ import annotations

import time

import pytest

from core import judgment, texts, workflow
from core.matching import Matcher
from core.models import Case, CaseRow, MasterItem, Rule
from core.normalize import norm, norm_wbs


def key_of(db, name):
    return next(i.key for i in workflow.master_items(db) if i.name == name)


# ---- 正規化 ----
@pytest.mark.parametrize("a, b", [("ＰＡＣＳ連携テスト", "PACS連携テスト"), ("現行業務 の 調査", "現行業務の調査"),
                                  ("【課題管理】", "課題管理"), ("定例会議（運営）", "定例会議運営")])
def test_norm(a, b):
    assert norm(a) == norm(b)


@pytest.mark.parametrize("raw, expected", [("2-1", "2.1"), ("２．１", "2.1"), ("02.01", "2.1"), ("1.0", "1"), (" 3 ", "3")])
def test_norm_wbs(raw, expected):
    assert norm_wbs(raw) == expected


# ---- 照合 ----
def test_matching_by_wbs_name_and_fuzzy(db, loaded):
    rows = db.case_rows()
    by_case = {}
    for r in rows:
        by_case.setdefault(r.case_id, []).append(r)
    a, b, c = (by_case[loaded[n]] for n in ("case_A", "case_B", "case_C"))
    assert all(r.status == "auto" and r.method == "wbs" for r in a), "A は WBS 番号で全件一致"
    # B：番号を振り直している（2-1 形式）。番号が一致しても名称がかけ離れていれば番号では確定しない
    b_status = {r.name: (r.status, r.method, r.item_key) for r in b}
    assert b_status["現行業務調査"][0] == "auto"                     # 類似度が高い表記ゆれ
    assert b_status["要件定義書の承認"][1] == "exact"               # 番号 2-4 はひな形の 2.4（帳票要件）と別物
    assert b_status["総合テスト"][2] == key_of(db, "総合テスト")     # 番号 4-5 は PACS連携テストと誤照合しない
    assert b_status["職員操作研修会"][0] == "pending", "類似度が中間帯なので確認待ち"
    cand_names = [key for key, _ in next(r for r in b if r.name == "職員操作研修会").candidates]
    assert key_of(db, "職員向け操作研修") == cand_names[0]
    # C：ひな形にない追加項目
    extra = [r for r in c if r.status == "extra"]
    assert [r.name for r in extra] == ["院内説明会の実施"]
    assert extra[0].prev_key == key_of(db, "職員向け操作研修")


def test_confirmed_choice_is_saved_as_synonym(db, cfg, loaded):
    row = next(r for r in db.case_rows() if r.name == "職員操作研修会")
    workflow.confirm_match(db, row.id, key_of(db, "職員向け操作研修"))
    assert db.synonyms()[norm("職員操作研修会")] == key_of(db, "職員向け操作研修")
    # 同じ表記の新しい事例は同義語辞書で自動確定する
    m = Matcher(workflow.master_items(db), db.synonyms(), cfg["match"])
    r = m.match_row(CaseRow(None, None, 0, 0, "", "教育・運用準備", "職員操作研修会", norm("職員操作研修会")))
    assert (r.status, r.method, r.item_key) == ("auto", "synonym", key_of(db, "職員向け操作研修"))


# ---- 採用率・閾値 ----
def test_adoption_rate_and_thresholds(db, cfg, loaded):
    cfg["weighting"]["enabled"] = False
    js, _ = workflow.run_judgment(db, cfg)
    by = {j.name: j for j in js}
    assert by["プロジェクト計画書の作成"].rate == 1.0 and by["プロジェクト計画書の作成"].decision == "必要"
    assert by["PACS連携テスト"].rate == pytest.approx(2 / 3) and by["PACS連携テスト"].decision == "要検討"
    assert by["DPC要件の定義"].adopted_cases == ["case_A"] and by["DPC要件の定義"].not_adopted_cases == ["case_B", "case_C"]
    assert by["DPC要件の定義"].decision == "要検討", "1/3 = 33% は 30% 以上なので要検討"
    assert by["プロジェクト計画書の作成"].checked and not by["PACS連携テスト"].checked, "初期チェックは「必要」だけ"
    extra = [j for j in js if j.is_extra]
    assert [e.name for e in extra] == ["院内説明会の実施"] and extra[0].rate == pytest.approx(1 / 3)
    cfg["thresholds"] = {"required": 0.6, "unneeded": 0.4}
    js2, _ = workflow.run_judgment(db, cfg)
    by2 = {j.name: j for j in js2}
    assert by2["PACS連携テスト"].decision == "必要" and by2["DPC要件の定義"].decision == "不要"


def test_parent_adopted_when_child_adopted():
    items = [MasterItem("p", 0, 1, "1", "P", "親", "親", 1, None), MasterItem("c", 1, 2, "1.1", "P", "子", "子", 2, "p")]
    cases = [Case(1, "X")]
    rows = [CaseRow(1, 1, 0, 0, "1.1", "P", "子", "子", item_key="c", status="auto")]
    js = judgment.judge(items, cases, rows, [], {}, {}, {"thresholds": {"required": 0.7, "unneeded": 0.3}})
    assert [j.rate for j in js] == [1.0, 1.0]


def test_weighting_by_project_attributes(db, cfg, loaded):
    db.save_project_attrs({"update_type": "リプレース", "beds": 450, "vendor": "ベンダーX", "systems": ["電子カルテ", "PACS"]})
    js, _ = workflow.run_judgment(db, cfg)
    pacs = next(j for j in js if j.name == "PACS連携テスト")
    assert pacs.wrate > pacs.rate, "PACS を使う近い事例（A・C）の重みが上がる"
    assert pacs.decision == "必要"
    cfg["weighting"]["enabled"] = False
    js, _ = workflow.run_judgment(db, cfg)
    assert next(j for j in js if j.name == "PACS連携テスト").wrate is None


def test_user_rule_overrides_statistics_and_manual_overrides_rule(db, cfg, loaded):
    cfg["weighting"]["enabled"] = False
    db.save_project_attrs({"systems": ["電子カルテ", "PACS"]})
    pacs = key_of(db, "PACS連携テスト")
    db.save_rule(Rule(None, "PACS連携あり", "PACS integration", True, 10,
                      [{"attr": "systems", "op": "含む", "value": "PACS"}], "必要", [pacs]))
    db.save_rule(Rule(None, "低優先", "", True, 20, [{"attr": "systems", "op": "含む", "value": "PACS"}], "不要", [pacs]))
    js, _ = workflow.run_judgment(db, cfg)
    j = next(j for j in js if j.key == pacs)
    assert j.stat_decision == "要検討" and j.decision == "必要" and j.rule_name == "PACS連携あり", "優先順位の高いルール"
    assert "条件『PACS連携あり』に一致。" in j.reason_ja and "Matches condition: PACS integration." in j.reason_en
    # 条件に合わなければ適用しない
    db.save_project_attrs({"systems": ["電子カルテ"]})
    js, _ = workflow.run_judgment(db, cfg)
    assert next(j for j in js if j.key == pacs).rule_name is None
    # 手動上書きはルールより優先。履歴が残る
    db.save_project_attrs({"systems": ["PACS"]})
    db.add_override(pacs, "不要", "必要", "今回は対象外")
    js, _ = workflow.run_judgment(db, cfg)
    j = next(j for j in js if j.key == pacs)
    assert j.decision == "不要" and not j.checked and "手動で「不要」に変更。" in j.reason_ja
    assert db.override_history(pacs)[0]["previous"] == "必要"
    db.add_override(pacs, None, "不要", "解除")
    js, _ = workflow.run_judgment(db, cfg)
    assert next(j for j in js if j.key == pacs).decision == "必要"


@pytest.mark.parametrize("cond, project, expected", [
    ({"attr": "beds", "op": "以上", "value": "400"}, {"beds": 450}, True),
    ({"attr": "beds", "op": "以下", "value": "400"}, {"beds": 450}, False),
    ({"attr": "update_type", "op": "＝", "value": "リプレース"}, {"update_type": "リプレース"}, True),
    ({"attr": "vendor", "op": "≠", "value": "X"}, {"vendor": "X"}, False),
    ({"attr": "systems", "op": "含まない", "value": "PACS"}, {"systems": ["電子カルテ"]}, True),
    ({"attr": "tags", "op": "未入力", "value": ""}, {}, True),
])
def test_rule_operators(cond, project, expected):
    assert judgment.condition_holds(cond, project) is expected


# ---- 日英の文章 ----
def test_reason_templates_ja_en():
    from core import config
    t = config.load_text_templates()
    j = judgment.Judgment(key="k", name="PACS連携テスト", wbs_no="4.5", phase="", level=2, is_extra=False, total=8,
                          adopted_cases=[f"c{i}" for i in range(7)], not_adopted_cases=["c7"], rate=7 / 8, wrate=None,
                          stat_decision="必要", rule_name="PACS連携あり", rule_name_en="PACS integration")
    texts.build_texts([j], {"PACS連携テスト": "PACS interface test"}, t)
    assert j.reason_ja == "過去8件中7件（88%）で採用。条件『PACS連携あり』に一致。"
    assert j.reason_en == "Adopted in 7 of 8 past projects (88%). Matches condition: PACS integration."
    assert j.name_en == "PACS interface test" and j.untranslated == []


def test_untranslated_terms_detected():
    tr = texts.Translator({"操作研修": "operation training", "PACS": "PACS", "職員向け操作研修": ""})
    assert tr.translate("職員向け操作研修") == ("【未翻訳: 職員向け】 operation training", ["職員向け"])
    assert tr.translate("稼働判定会議") == ("【未翻訳: 稼働判定会議】", ["稼働判定会議"])
    assert tr.translate("PACS連携")[1] == ["連携"]


def test_glossary_initialized_with_master_items(db, loaded):
    g = db.glossary()
    assert "PACS連携テスト" in g and g["PACS連携テスト"] == ""
    assert len([k for k in g]) >= 44


def test_untranslated_count_and_copy_formats(db, cfg, loaded):
    db.upsert_terms({"プロジェクト計画書の作成": "Project plan preparation"}, keep_existing_en=False)
    js, untranslated = workflow.run_judgment(db, cfg)
    j = next(j for j in js if j.name == "プロジェクト計画書の作成")
    assert j.name_en == "Project plan preparation" and not j.untranslated
    assert "PACS連携テスト" in untranslated
    from core import config
    t = config.load_text_templates()
    tsv_ja = texts.copy_text([j], "ja", "tsv", t)
    tsv_en = texts.copy_text([j], "en", "tsv", t)
    assert tsv_ja.split("\n")[0] == "WBS番号\t作業項目\t判定\t採用率\t判定理由"
    assert tsv_ja.split("\n")[1].split("\t")[:4] == ["1.1", "プロジェクト計画書の作成", "必要", "100%"]
    assert tsv_en.split("\n")[1].split("\t")[:3] == ["1.1", "Project plan preparation", "Required"]
    txt = texts.copy_text([j], "en", "text", t)
    assert txt.startswith("1.1 Project plan preparation: Required (100%) Adopted in 3 of 3")
    assert "\t" not in txt


# ---- 性能：ひな形 1000 行・過去事例 50 件 ----
def test_performance_1000_items_50_cases():
    items, rows, cases = [], [], []
    for i in range(1000):
        items.append(MasterItem(f"k{i}", i, i, f"{i // 50 + 1}.{i % 50 + 1}", f"工程{i // 50}", f"作業項目{i:04d}の実施",
                                norm(f"作業項目{i:04d}の実施"), 2))
    rid = 0
    for c in range(50):
        cases.append(Case(c, f"事例{c}", attrs={"beds": 100 + c * 10}))
        for i in range(0, 1000, 2 + c % 3):
            rid += 1
            name = f"作業項目{i:04d}の実施" if i % 7 else f"作業項目{i:04d}を実施"
            rows.append(CaseRow(rid, c, i, i, "", f"工程{i // 50}", name, norm(name)))
    cfg = {"match": {"auto": 90, "review": 70}, "thresholds": {"required": 0.7, "unneeded": 0.3},
           "weighting": {"enabled": True}}
    t0 = time.perf_counter()
    m = Matcher(items, {}, cfg["match"])
    by_case = {}
    for r in rows:
        by_case.setdefault(r.case_id, []).append(r)
    matched = []
    for rs in by_case.values():
        for x, res in zip(sorted(rs, key=lambda r: r.seq), m.match_rows(rs)):
            x.item_key, x.status = res.item_key, res.status
            matched.append(x)
    js = judgment.judge(items, cases, matched, [], {"beds": 300}, {}, cfg)
    texts.build_texts(js, {}, {"stat": {"ja": "", "en": ""}, "stat_none": {"ja": "", "en": ""}, "weighted": {"ja": "", "en": ""},
                                "no_cases": {"ja": "", "en": ""}, "rule": {"ja": "", "en": ""}, "manual": {"ja": "", "en": ""},
                                "extra": {"ja": "", "en": ""}, "decisions": {}})
    assert time.perf_counter() - t0 < 10
    assert len(js) >= 1000
