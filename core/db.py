"""SQLite（%LOCALAPPDATA%\\WBSAdvisor\\wbs.db）。

WBS の種類（設計・構築・テスト など）ごとに、ひな形・実用WBS様式・過去事例・同義語辞書・ルール・手動上書きを分けて持つ。
Database.wbs_type に現在の種類を設定すると、以降の読み書きはその種類だけが対象になる。
今回の案件の属性と用語辞書は種類をまたいで共通。

テーブル
  formats       登録した様式（kind: master=ひな形 / output=実用WBS様式）。profile・analysis は JSON
  master_items  ひな形の項目（item_key はひな形を登録し直しても変わらない識別子）
  cases         過去事例（sha256 で二重取込を検出。種類ごと）。attrs は案件属性の JSON
  case_rows     過去事例の各行と、ひな形項目への対応付け（method / score / status）
  synonyms      確認画面で選んだ対応付け（正規化した名称 → item_key）。次回以降に再利用
  rules         ユーザー独自条件
  project       今回の案件の属性（1 行だけ。種類をまたいで共通）
  overrides     判定の手動上書きの履歴（item_key ごとに最新の 1 件が有効。decision が NULL なら解除）
  glossary      用語辞書（日本語 → 英語。種類をまたいで共通）
"""
from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

from . import paths

DEFAULT_TYPE = "設計"   # 種類の列がない古いデータベースの行は、この種類として扱う

SCHEMA = """
CREATE TABLE IF NOT EXISTS formats(
  id INTEGER PRIMARY KEY, wbs_type TEXT NOT NULL DEFAULT '設計',
  kind TEXT NOT NULL CHECK(kind IN ('master','output')), name TEXT NOT NULL,
  file TEXT NOT NULL, source TEXT, sha256 TEXT, profile TEXT, analysis TEXT, created TEXT, active INTEGER DEFAULT 1);
CREATE TABLE IF NOT EXISTS master_items(
  id INTEGER PRIMARY KEY, format_id INTEGER NOT NULL, item_key TEXT NOT NULL, seq INTEGER, row INTEGER,
  wbs_no TEXT, phase TEXT, name TEXT, norm TEXT, level INTEGER, parent_key TEXT, values_json TEXT);
CREATE INDEX IF NOT EXISTS ix_master_format ON master_items(format_id);
CREATE TABLE IF NOT EXISTS cases(
  id INTEGER PRIMARY KEY, wbs_type TEXT NOT NULL DEFAULT '設計', name TEXT NOT NULL, source TEXT, sha256 TEXT,
  imported TEXT, attrs TEXT);
CREATE TABLE IF NOT EXISTS case_rows(
  id INTEGER PRIMARY KEY, case_id INTEGER NOT NULL REFERENCES cases(id) ON DELETE CASCADE, seq INTEGER, row INTEGER,
  wbs_no TEXT, phase TEXT, name TEXT, norm TEXT, values_json TEXT,
  item_key TEXT, method TEXT, score REAL, status TEXT, candidates TEXT, prev_key TEXT);
CREATE INDEX IF NOT EXISTS ix_rows_case ON case_rows(case_id);
CREATE TABLE IF NOT EXISTS synonyms(
  wbs_type TEXT NOT NULL DEFAULT '設計', norm TEXT NOT NULL, item_key TEXT NOT NULL, created TEXT,
  PRIMARY KEY(wbs_type, norm));
CREATE TABLE IF NOT EXISTS rules(
  id INTEGER PRIMARY KEY, wbs_type TEXT NOT NULL DEFAULT '設計', name TEXT NOT NULL, name_en TEXT,
  enabled INTEGER DEFAULT 1, priority INTEGER DEFAULT 100, conditions TEXT, decision TEXT, targets TEXT, keyword TEXT,
  created TEXT);
CREATE TABLE IF NOT EXISTS project(id INTEGER PRIMARY KEY CHECK(id = 1), attrs TEXT);
CREATE TABLE IF NOT EXISTS overrides(
  id INTEGER PRIMARY KEY, wbs_type TEXT NOT NULL DEFAULT '設計', item_key TEXT NOT NULL, decision TEXT, previous TEXT,
  note TEXT, at TEXT);
CREATE TABLE IF NOT EXISTS glossary(ja TEXT PRIMARY KEY, en TEXT DEFAULT '', source TEXT);
"""


def now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _columns(conn, table) -> list[str]:
    return [r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]


def _migrate(conn) -> None:
    """種類の列がない古いデータベースを移行する（既存の行は DEFAULT_TYPE にする）。"""
    tables = {r[0]: r[1] for r in conn.execute("SELECT name, sql FROM sqlite_master WHERE type = 'table'").fetchall()}
    if not tables:
        return
    conn.execute("PRAGMA foreign_keys = OFF")
    try:
        for t in ("formats", "rules", "overrides"):
            if t in tables and "wbs_type" not in _columns(conn, t):
                conn.execute(f"ALTER TABLE {t} ADD COLUMN wbs_type TEXT NOT NULL DEFAULT '{DEFAULT_TYPE}'")
        # 作り直しは SQLite の推奨手順（新しい表を作って写す → 古い表を消す → 新しい表の名前を戻す）で行う。
        # 古い表の名前を先に変えると、case_rows の外部キーの参照先まで書き換わってしまうため。
        # cases：sha256 の UNIQUE をやめて、種類ごとに二重取込を判定する
        if "cases" in tables and "wbs_type" not in _columns(conn, "cases"):
            conn.execute(f"CREATE TABLE cases_new(id INTEGER PRIMARY KEY, wbs_type TEXT NOT NULL DEFAULT '{DEFAULT_TYPE}', "
                         "name TEXT NOT NULL, source TEXT, sha256 TEXT, imported TEXT, attrs TEXT)")
            conn.execute("INSERT INTO cases_new(id, name, source, sha256, imported, attrs) "
                         "SELECT id, name, source, sha256, imported, attrs FROM cases")
            conn.execute("DROP TABLE cases")
            conn.execute("ALTER TABLE cases_new RENAME TO cases")
        # synonyms：主キーを (種類, 名称) にする
        if "synonyms" in tables and "wbs_type" not in _columns(conn, "synonyms"):
            conn.execute(f"CREATE TABLE synonyms_new(wbs_type TEXT NOT NULL DEFAULT '{DEFAULT_TYPE}', norm TEXT NOT NULL, "
                         "item_key TEXT NOT NULL, created TEXT, PRIMARY KEY(wbs_type, norm))")
            conn.execute("INSERT INTO synonyms_new(norm, item_key, created) SELECT norm, item_key, created FROM synonyms")
            conn.execute("DROP TABLE synonyms")
            conn.execute("ALTER TABLE synonyms_new RENAME TO synonyms")
        conn.commit()
    finally:
        conn.execute("PRAGMA foreign_keys = ON")


class Database:
    def __init__(self, path: Path | None = None, wbs_type: str = DEFAULT_TYPE):
        self.path = Path(path) if path else paths.db_path()
        self.conn = sqlite3.connect(self.path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        _migrate(self.conn)
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.executescript(SCHEMA)
        self.conn.commit()
        self.wbs_type = wbs_type

    @contextmanager
    def tx(self):
        try:
            yield self.conn
            self.conn.commit()
        except BaseException:
            self.conn.rollback()
            raise

    def close(self):
        self.conn.close()

    def counts_by_type(self) -> dict[str, dict[str, int]]:
        """種類ごとの登録状況（画面の表示用）。"""
        out: dict[str, dict[str, int]] = {}
        for r in self.conn.execute("SELECT wbs_type, kind, COUNT(*) n FROM formats WHERE active = 1 GROUP BY wbs_type, kind"):
            out.setdefault(r["wbs_type"], {})[r["kind"]] = r["n"]
        for r in self.conn.execute("SELECT wbs_type, COUNT(*) n FROM cases GROUP BY wbs_type"):
            out.setdefault(r["wbs_type"], {})["cases"] = r["n"]
        return out

    # ---- 様式 ----
    def add_format(self, kind, name, file, source, sha256, profile, analysis) -> int:
        with self.tx() as c:
            c.execute("UPDATE formats SET active = 0 WHERE kind = ? AND wbs_type = ?", (kind, self.wbs_type))
            cur = c.execute(
                "INSERT INTO formats(wbs_type, kind, name, file, source, sha256, profile, analysis, created, active) "
                "VALUES(?,?,?,?,?,?,?,?,?,1)",
                (self.wbs_type, kind, name, str(file), str(source), sha256, json.dumps(profile, ensure_ascii=False),
                 json.dumps(analysis, ensure_ascii=False), now()))
            return cur.lastrowid

    def active_format(self, kind) -> dict | None:
        r = self.conn.execute("SELECT * FROM formats WHERE kind = ? AND wbs_type = ? AND active = 1 "
                              "ORDER BY id DESC LIMIT 1", (kind, self.wbs_type)).fetchone()
        if not r:
            return None
        d = dict(r)
        d["profile"] = json.loads(d["profile"] or "{}")
        d["analysis"] = json.loads(d["analysis"] or "{}")
        return d

    def update_format_profile(self, format_id, profile, analysis=None):
        with self.tx() as c:
            c.execute("UPDATE formats SET profile = ? WHERE id = ?", (json.dumps(profile, ensure_ascii=False), format_id))
            if analysis is not None:
                c.execute("UPDATE formats SET analysis = ? WHERE id = ?",
                          (json.dumps(analysis, ensure_ascii=False), format_id))

    # ---- ひな形の項目 ----
    def replace_master_items(self, format_id, items):
        with self.tx() as c:
            c.execute("DELETE FROM master_items WHERE format_id = ?", (format_id,))
            c.executemany(
                "INSERT INTO master_items(format_id, item_key, seq, row, wbs_no, phase, name, norm, level, parent_key, "
                "values_json) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                [(format_id, i.key, i.seq, i.row, i.wbs_no, i.phase, i.name, i.norm, i.level, i.parent_key,
                  json.dumps(i.values, ensure_ascii=False, default=str)) for i in items])

    def master_items(self, format_id):
        from .models import MasterItem
        rows = self.conn.execute("SELECT * FROM master_items WHERE format_id = ? ORDER BY seq", (format_id,)).fetchall()
        return [MasterItem(key=r["item_key"], seq=r["seq"], row=r["row"], wbs_no=r["wbs_no"], phase=r["phase"],
                           name=r["name"], norm=r["norm"], level=r["level"], parent_key=r["parent_key"],
                           values=json.loads(r["values_json"] or "{}")) for r in rows]

    # ---- 過去事例 ----
    def case_by_hash(self, sha256):
        r = self.conn.execute("SELECT * FROM cases WHERE sha256 = ? AND wbs_type = ?", (sha256, self.wbs_type)).fetchone()
        return dict(r) if r else None

    def add_case(self, name, source, sha256, attrs, rows) -> int:
        with self.tx() as c:
            cur = c.execute("INSERT INTO cases(wbs_type, name, source, sha256, imported, attrs) VALUES(?,?,?,?,?,?)",
                            (self.wbs_type, name, str(source), sha256, now(), json.dumps(attrs or {}, ensure_ascii=False)))
            cid = cur.lastrowid
            c.executemany(
                "INSERT INTO case_rows(case_id, seq, row, wbs_no, phase, name, norm, values_json) VALUES(?,?,?,?,?,?,?,?)",
                [(cid, r.seq, r.row, r.wbs_no, r.phase, r.name, r.norm, json.dumps(r.values, ensure_ascii=False, default=str))
                 for r in rows])
            return cid

    def cases(self):
        from .models import Case
        out = []
        for r in self.conn.execute("SELECT * FROM cases WHERE wbs_type = ? ORDER BY id", (self.wbs_type,)).fetchall():
            out.append(Case(id=r["id"], name=r["name"], source=r["source"], sha256=r["sha256"],
                            imported=r["imported"], attrs=json.loads(r["attrs"] or "{}")))
        return out

    def update_case_attrs(self, case_id, attrs):
        with self.tx() as c:
            c.execute("UPDATE cases SET attrs = ? WHERE id = ?", (json.dumps(attrs, ensure_ascii=False), case_id))

    def rename_case(self, case_id, name):
        with self.tx() as c:
            c.execute("UPDATE cases SET name = ? WHERE id = ?", (name, case_id))

    def delete_case(self, case_id):
        with self.tx() as c:
            c.execute("DELETE FROM case_rows WHERE case_id = ?", (case_id,))
            c.execute("DELETE FROM cases WHERE id = ?", (case_id,))

    def case_rows(self, case_id=None):
        from .models import CaseRow
        if case_id is not None:
            rows = self.conn.execute("SELECT * FROM case_rows WHERE case_id = ? ORDER BY seq", (case_id,)).fetchall()
        else:
            rows = self.conn.execute("SELECT r.* FROM case_rows r JOIN cases c ON c.id = r.case_id WHERE c.wbs_type = ? "
                                     "ORDER BY r.case_id, r.seq", (self.wbs_type,)).fetchall()
        return [CaseRow(id=r["id"], case_id=r["case_id"], seq=r["seq"], row=r["row"], wbs_no=r["wbs_no"],
                        phase=r["phase"], name=r["name"], norm=r["norm"], values=json.loads(r["values_json"] or "{}"),
                        item_key=r["item_key"], method=r["method"], score=r["score"], status=r["status"],
                        candidates=json.loads(r["candidates"] or "[]"), prev_key=r["prev_key"]) for r in rows]

    def save_matches(self, results):
        with self.tx() as c:
            c.executemany(
                "UPDATE case_rows SET item_key=?, method=?, score=?, status=?, candidates=?, prev_key=? WHERE id=?",
                [(m.item_key, m.method, m.score, m.status, json.dumps(m.candidates, ensure_ascii=False), m.prev_key,
                  m.row_id) for m in results])

    def set_row_match(self, row_id, item_key, status, method="manual"):
        with self.tx() as c:
            c.execute("UPDATE case_rows SET item_key=?, status=?, method=? WHERE id=?", (item_key, status, method, row_id))

    # ---- 同義語辞書 ----
    def synonyms(self) -> dict[str, str]:
        return {r["norm"]: r["item_key"] for r in
                self.conn.execute("SELECT * FROM synonyms WHERE wbs_type = ?", (self.wbs_type,)).fetchall()}

    def add_synonym(self, norm, item_key):
        with self.tx() as c:
            c.execute("INSERT OR REPLACE INTO synonyms(wbs_type, norm, item_key, created) VALUES(?,?,?,?)",
                      (self.wbs_type, norm, item_key, now()))

    # ---- ルール ----
    def rules(self):
        from .models import Rule
        out = []
        for r in self.conn.execute("SELECT * FROM rules WHERE wbs_type = ? ORDER BY priority, id", (self.wbs_type,)).fetchall():
            out.append(Rule(id=r["id"], name=r["name"], name_en=r["name_en"] or "", enabled=bool(r["enabled"]),
                            priority=r["priority"], conditions=json.loads(r["conditions"] or "[]"),
                            decision=r["decision"], targets=json.loads(r["targets"] or "[]"), keyword=r["keyword"] or ""))
        return out

    def save_rule(self, rule) -> int:
        vals = (rule.name, rule.name_en, int(rule.enabled), rule.priority, json.dumps(rule.conditions, ensure_ascii=False),
                rule.decision, json.dumps(rule.targets, ensure_ascii=False), rule.keyword)
        with self.tx() as c:
            if rule.id:
                c.execute("UPDATE rules SET name=?, name_en=?, enabled=?, priority=?, conditions=?, decision=?, "
                          "targets=?, keyword=? WHERE id=?", vals + (rule.id,))
                return rule.id
            cur = c.execute("INSERT INTO rules(name, name_en, enabled, priority, conditions, decision, targets, keyword, "
                            "created, wbs_type) VALUES(?,?,?,?,?,?,?,?,?,?)", vals + (now(), self.wbs_type))
            return cur.lastrowid

    def delete_rule(self, rule_id):
        with self.tx() as c:
            c.execute("DELETE FROM rules WHERE id = ?", (rule_id,))

    # ---- 今回の案件（種類をまたいで共通） ----
    def project_attrs(self) -> dict:
        r = self.conn.execute("SELECT attrs FROM project WHERE id = 1").fetchone()
        return json.loads(r["attrs"]) if r and r["attrs"] else {}

    def save_project_attrs(self, attrs):
        with self.tx() as c:
            c.execute("INSERT OR REPLACE INTO project(id, attrs) VALUES(1, ?)", (json.dumps(attrs, ensure_ascii=False),))

    # ---- 手動上書き ----
    def add_override(self, item_key, decision, previous, note=""):
        with self.tx() as c:
            c.execute("INSERT INTO overrides(wbs_type, item_key, decision, previous, note, at) VALUES(?,?,?,?,?,?)",
                      (self.wbs_type, item_key, decision, previous, note, now()))

    def overrides(self) -> dict[str, str]:
        out = {}
        for r in self.conn.execute("SELECT item_key, decision FROM overrides WHERE wbs_type = ? ORDER BY id",
                                   (self.wbs_type,)).fetchall():
            if r["decision"]:
                out[r["item_key"]] = r["decision"]
            else:
                out.pop(r["item_key"], None)
        return out

    def override_history(self, item_key=None):
        q = "SELECT * FROM overrides WHERE wbs_type = ?" + (" AND item_key = ?" if item_key else "") + " ORDER BY id DESC"
        return [dict(r) for r in self.conn.execute(q, (self.wbs_type, item_key) if item_key else (self.wbs_type,)).fetchall()]

    # ---- 用語辞書（種類をまたいで共通） ----
    def glossary(self) -> dict[str, str]:
        return {r["ja"]: r["en"] or "" for r in self.conn.execute("SELECT ja, en FROM glossary").fetchall()}

    def upsert_terms(self, terms: dict[str, str], source="user", keep_existing_en=True):
        with self.tx() as c:
            for ja, en in terms.items():
                ja = (ja or "").strip()
                if not ja:
                    continue
                if keep_existing_en and not en:
                    c.execute("INSERT OR IGNORE INTO glossary(ja, en, source) VALUES(?, '', ?)", (ja, source))
                else:
                    c.execute("INSERT INTO glossary(ja, en, source) VALUES(?,?,?) "
                              "ON CONFLICT(ja) DO UPDATE SET en = excluded.en, source = excluded.source",
                              (ja, (en or "").strip(), source))

    def delete_term(self, ja):
        with self.tx() as c:
            c.execute("DELETE FROM glossary WHERE ja = ?", (ja,))
