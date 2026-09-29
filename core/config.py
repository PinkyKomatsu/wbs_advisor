"""設定（JSON）。初期値は resources/default_config.json、利用者の設定は %LOCALAPPDATA%\\WBSAdvisor\\config.json。"""
from __future__ import annotations

import copy
import json
import os
import tempfile
from pathlib import Path

from . import paths


def _deep_merge(base: dict, over: dict) -> dict:
    out = copy.deepcopy(base)
    for k, v in (over or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def _read(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}
    except (OSError, json.JSONDecodeError):
        return {}


def save_json(path: Path, obj) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=path.parent, suffix=".tmp")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def defaults() -> dict:
    return _read(paths.resource("default_config.json"))


def load() -> dict:
    return _deep_merge(defaults(), _read(paths.config_path()))


def save(cfg: dict) -> None:
    save_json(paths.config_path(), cfg)


def load_text_templates() -> dict:
    return _deep_merge(_read(paths.resource("text_templates.json")), _read(paths.templates_path()))


def save_text_templates(t: dict) -> None:
    save_json(paths.templates_path(), t)
