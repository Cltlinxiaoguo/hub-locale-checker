# -*- coding: utf-8 -*-
"""从 config.json 读取路径与输出选项。"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
DEFAULT_CONFIG = ROOT / "config.json"

# 默认子目录（config 未写时使用）
DEFAULT_JS_DIR = "new file"
DEFAULT_RULE_DIR = "rule file"
DEFAULT_OUTPUT_DIR = "report"


def load_config(config_path: Path | None = None) -> dict[str, Any]:
    path = config_path or DEFAULT_CONFIG
    if not path.is_file():
        raise SystemExit(
            f"未找到配置文件: {path}\n请复制 config.json 并填写 paths 等项。"
        )
    with path.open(encoding="utf-8") as f:
        data = json.load(f)
    return data


def _resolve(base: Path, p: str) -> Path:
    path = Path(p)
    return path if path.is_absolute() else (base / path)


def _ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def get_paths(cfg: dict[str, Any]) -> dict[str, Path]:
    paths = cfg.get("paths") or {}

    js_dir = _ensure_dir(_resolve(ROOT, paths.get("js_dir") or DEFAULT_JS_DIR))
    rule_dir = _ensure_dir(_resolve(ROOT, paths.get("rule_dir") or DEFAULT_RULE_DIR))
    out_dir = _ensure_dir(_resolve(ROOT, paths.get("output_dir") or DEFAULT_OUTPUT_DIR))

    js_name = paths.get("js_file")
    if not js_name:
        candidates = sorted(
            js_dir.glob("language-*.js"), key=lambda x: x.stat().st_mtime, reverse=True
        )
        if not candidates:
            raise SystemExit(
                f"config.paths.js_file 为空，且 {js_dir} 下无 language-*.js"
            )
        js_path = candidates[0]
    else:
        p = Path(js_name)
        js_path = p if p.is_absolute() else _resolve(js_dir, js_name)
        if not js_path.is_file():
            raise SystemExit(f"JS 文件不存在: {js_path}\n请放入目录: {js_dir}")

    std_name = paths.get("standard_xlsx")
    std_path = None
    if std_name:
        p = Path(std_name)
        std_path = p if p.is_absolute() else _resolve(rule_dir, std_name)

    return {
        "js_dir": js_dir,
        "rule_dir": rule_dir,
        "output_dir": out_dir,
        "js_file": js_path,
        "standard_xlsx": std_path,
    }


def resolve_js_override(js_cli: str, cfg: dict[str, Any]) -> Path:
    """命令行传入的 JS 路径：相对路径优先在 js_dir 下查找。"""
    p = Path(js_cli)
    if p.is_absolute():
        return p
    paths = get_paths(cfg)
    candidate = paths["js_dir"] / p
    if candidate.is_file():
        return candidate
    root_candidate = ROOT / p
    if root_candidate.is_file():
        return root_candidate
    return candidate


def timestamp_suffix(cfg: dict[str, Any]) -> str:
    out = cfg.get("output") or {}
    if not out.get("xlsx_add_timestamp", True):
        return ""
    fmt = out.get("timestamp_format") or "%Y%m%d_%H%M%S"
    return datetime.now().strftime(fmt)


def output_json_path(js_path: Path, cfg: dict[str, Any]) -> Path:
    d = get_paths(cfg)["output_dir"]
    return d / f"{js_path.stem}_zh_en.json"


def output_report_path(js_path: Path, cfg: dict[str, Any]) -> Path:
    d = get_paths(cfg)["output_dir"]
    return d / f"{js_path.stem}_compare_report.txt"


def output_xlsx_path(js_path: Path, cfg: dict[str, Any]) -> Path:
    d = get_paths(cfg)["output_dir"]
    ts = timestamp_suffix(cfg)
    name = f"{js_path.stem}_中英对照"
    if ts:
        name = f"{name}_{ts}"
    return d / f"{name}.xlsx"


def compare_enabled(cfg: dict[str, Any]) -> bool:
    return (cfg.get("compare") or {}).get("enabled", True)
