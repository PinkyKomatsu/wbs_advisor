"""画面から使う一連の処理（様式登録・事例取込・照合・判定・出力）。GUI には依存しない。"""
from __future__ import annotations

import logging
import shutil
from pathlib import Path

from excel_io import engine as engine_mod
from excel_io import inspect, plan as plan_mod, profile as profile_mod, reader, verify as verify_mod

from . import config as config_mod, paths, texts
from .db import Database
from .judgment import judge, output_order
from .matching import Matcher
from .models import CaseRow, MasterItem
from .normalize import item_key, norm, wbs_level

log = logging.getLogger("wbsadvisor")


class DuplicateCase(Exception):
    pass


# ---------------------------------------------------------------------------
# 様式の登録
# ---------------------------------------------------------------------------

def _copy_in(path: Path, kind: str) -> Path:
    dest_dir = paths.formats_dir() / kind
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / path.name
    if dest.resolve() != path.resolve():
        shutil.copy2(path, dest)
    return dest


def build_master_items(rows: list[tuple[int, dict]]) -> list[MasterItem]:
    """ひな形の明細からひな形項目を作る（階層は WBS 番号から。親は直前の浅い階層の項目）。"""
    items: list[MasterItem] = []
    keys: set[str] = set()
    current_phase = ""
    for seq, (r, v) in enumerate(rows):
        name = reader.cell_text(v.get("name"))
        if not name:
            continue
        wbs = reader.cell_text(v.get("wbs_no"))
        level = wbs_level(wbs) if wbs else (items[-1].level if items else 1)
        phase = reader.cell_text(v.get("phase"))
        if level == 1 and not phase:
            current_phase = name
        phase = phase or current_phase
        key = item_key(phase, name)
        if key in keys:
            key = f"{key}#{wbs or seq}"
        keys.add(key)
        parent = next((i.key for i in reversed(items) if i.level < level), None)
        values = {k: (val.isoformat() if hasattr(val, "isoformat") else val) for k, val in v.items()}
        items.append(MasterItem(key=key, seq=seq, row=r, wbs_no=wbs, phase=phase, name=name, norm=norm(name),
                                level=level, parent_key=parent, values=values))
    return items


def register_master(db: Database, path, cfg: dict, profile: dict | None = None) -> dict:
    p = reader.check(path)
    stored = _copy_in(p, "master")
    prof = profile or profile_mod.detect(stored, "master", cfg)
    fid = db.add_format("master", p.stem, stored, p, reader.sha256(p), prof, {})
    items = build_master_items(profile_mod.read_items(stored, prof, cfg.get("profile", {}).get("blank_rows_to_stop", 5)))
    db.replace_master_items(fid, items)
    # 用語辞書の初期値：ひな形の全項目（英語は空欄）
    terms = {i.name: "" for i in items}
    terms.update({i.phase: "" for i in items if i.phase})
    db.upsert_terms(terms, source="ひな形")
    log.info("ひな形を登録しました（%d項目）", len(items))
    return {"format_id": fid, "profile": prof, "items": len(items)}


def reload_master(db: Database, cfg: dict, profile: dict) -> int:
    f = db.active_format("master")
    db.update_format_profile(f["id"], profile)
    items = build_master_items(profile_mod.read_items(f["file"], profile))
    db.replace_master_items(f["id"], items)
    db.upsert_terms({i.name: "" for i in items}, source="ひな形")
    return len(items)


def register_output(db: Database, path, cfg: dict, profile: dict | None = None) -> dict:
    p = reader.check(path)
    stored = _copy_in(p, "output")
    prof = profile or profile_mod.detect(stored, "output", cfg)
    analysis = inspect.analyze(stored, prof["sheet"])
    fid = db.add_format("output", p.stem, stored, p, reader.sha256(p), prof, analysis)
    log.info("実用WBS様式を登録しました")
    return {"format_id": fid, "profile": prof, "analysis": analysis}


def update_output_profile(db: Database, profile: dict) -> dict:
    f = db.active_format("output")
    analysis = inspect.analyze(f["file"], profile["sheet"])
    db.update_format_profile(f["id"], profile, analysis)
    return analysis


def master_items(db: Database) -> list[MasterItem]:
    f = db.active_format("master")
    return db.master_items(f["id"]) if f else []


# ---------------------------------------------------------------------------
# 過去事例
# ---------------------------------------------------------------------------

def case_profile(db: Database, path, cfg: dict) -> dict:
    """過去事例は登録済みの実用WBS様式のプロファイルで読む（シート・列が合わなければ自動推定）。"""
    out = db.active_format("output")
    prof = dict(out["profile"]) if out else None
    if prof:
        prof["data_end"] = None      # 過去事例は明細が長くてもよい
        wb = reader.load(path, read_only=True)
        ok = prof["sheet"] in wb.sheetnames
        wb.close()
        if ok:
            return prof
    detected = profile_mod.detect(path, "master", cfg)
    detected["data_end"] = None
    return detected


def import_case(db: Database, path, cfg: dict, attrs: dict | None = None, name: str | None = None,
                allow_duplicate: bool = False) -> int:
    p = reader.check(path)
    h = reader.sha256(p)
    dup = db.case_by_hash(h)
    if dup and not allow_duplicate:
        raise DuplicateCase(f"「{p.name}」は取込済みです（事例「{dup['name']}」と同じファイル）。")
    prof = case_profile(db, p, cfg)
    rows = []
    for seq, (r, v) in enumerate(profile_mod.read_items(p, prof, cfg.get("profile", {}).get("blank_rows_to_stop", 5))):
        nm = reader.cell_text(v.get("name"))
        if not nm:
            continue
        values = {k: (val.isoformat() if hasattr(val, "isoformat") else val) for k, val in v.items()}
        rows.append(CaseRow(id=None, case_id=None, seq=seq, row=r, wbs_no=reader.cell_text(v.get("wbs_no")),
                            phase=reader.cell_text(v.get("phase")), name=nm, norm=norm(nm), values=values))
    cid = db.add_case(name or p.stem, p, h, attrs or {}, rows)
    rematch(db, cfg, case_id=cid)
    log.info("過去事例を取り込みました（%d行）", len(rows))
    return cid


def rematch(db: Database, cfg: dict, case_id: int | None = None) -> None:
    """照合し直す（確認画面で確定した行はそのまま）。"""
    items = master_items(db)
    matcher = Matcher(items, db.synonyms(), cfg.get("match", {}))
    rows = db.case_rows(case_id)
    by_case: dict[int, list[CaseRow]] = {}
    for r in rows:
        by_case.setdefault(r.case_id, []).append(r)
    results = []
    for case_rows in by_case.values():
        fixed = {r.id: r for r in case_rows if r.status == "confirmed" or r.method == "manual"}
        for m in matcher.match_rows(case_rows):
            if m.row_id in fixed:
                continue
            results.append(m)
    db.save_matches(results)


def confirm_match(db: Database, row_id: int, item_key_: str | None, remember: bool = True) -> None:
    """確認画面での選択。item_key_ が None なら「ひな形外の追加項目」として確定。"""
    row = next(r for r in db.case_rows() if r.id == row_id)
    if item_key_:
        db.set_row_match(row_id, item_key_, "confirmed")
        if remember:
            db.add_synonym(row.norm, item_key_)
    else:
        db.set_row_match(row_id, None, "extra")


# ---------------------------------------------------------------------------
# 判定
# ---------------------------------------------------------------------------

def run_judgment(db: Database, cfg: dict, templates: dict | None = None):
    templates = templates or config_mod.load_text_templates()
    items = master_items(db)
    judgments = judge(items, db.cases(), db.case_rows(), db.rules(), db.project_attrs(), db.overrides(), cfg)
    untranslated = texts.build_texts(judgments, db.glossary(), templates)
    return judgments, untranslated


# ---------------------------------------------------------------------------
# 出力
# ---------------------------------------------------------------------------

def output_rows(judgments, include_parents: bool = True) -> list[dict]:
    rows = []
    for j in output_order(judgments, include_parents):
        v = dict(j.values or {})
        v.update({"wbs_no": "" if j.is_extra else j.wbs_no, "phase": j.phase, "name": j.name})
        rows.append(v)
    return rows


def build_plan(db: Database, cfg: dict, judgments, sources: dict | None = None, engine: str | None = None):
    out = db.active_format("output")
    master = db.active_format("master")
    if not out:
        raise RuntimeError("実用WBS様式（出力テンプレート）が登録されていません。")
    prof, analysis = out["profile"], out["analysis"]
    master_meanings = set((master or {}).get("profile", {}).get("columns", {}).values())
    sources = sources or prof.get("sources") or plan_mod.default_sources(prof, master_meanings, analysis)
    eng = engine or engine_mod.choose(cfg.get("output", {}).get("engine", "auto"))
    rows = output_rows(judgments, cfg.get("output", {}).get("include_parents", True))
    cols = list(prof.get("columns", {}).keys())
    current = plan_mod.template_values(out["file"], prof["sheet"], prof["data_start"], prof.get("data_end") or prof["data_start"],
                                       cols) if prof.get("data_start") else {}
    return plan_mod.build_plan(analysis, prof, rows, sources, cfg.get("value_map", {}), eng,
                               cfg.get("output", {}).get("clear_sample_rows", True), current)


def export(db: Database, plan, out_dir: Path | None = None):
    out = db.active_format("output")
    template = Path(out["file"])
    path = engine_mod.execute(template, out_dir or paths.output_dir(), plan)
    checks = verify_mod.verify(template, path, plan, out["analysis"])
    log.info("出力しました（%s方式、%dセル）", plan.engine, len(plan.actions))
    return path, checks
