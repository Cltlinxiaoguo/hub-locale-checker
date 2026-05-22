# -*- coding: utf-8 -*-
"""从 config.json 读取路径，提取 JS 中英词条并与标准 xlsx 对比。"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from i18n_config import (
    ROOT,
    compare_enabled,
    get_paths,
    load_config,
    output_json_path,
    output_report_path,
    resolve_js_override,
)


def extract_balanced_object(source: str, marker: str) -> str:
    idx = source.find(marker)
    if idx == -1:
        raise SystemExit(f"marker not found: {marker!r}")
    brace_start = source.find("{", idx)
    if brace_start == -1:
        raise SystemExit("no { after marker")
    depth = 0
    i = brace_start
    in_string = False
    string_quote = ""
    escape = False
    n = len(source)
    while i < n:
        c = source[i]
        if in_string:
            if escape:
                escape = False
            elif c == "\\":
                escape = True
            elif c == string_quote:
                in_string = False
                string_quote = ""
            i += 1
            continue
        if c in "\"'`":
            in_string = True
            string_quote = c
            i += 1
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return source[brace_start : i + 1]
        i += 1
    raise SystemExit("unbalanced braces")


def parse_js_object_literal(literal: str) -> dict:
    import subprocess

    fixed = literal.replace("vxe:ce.vxe", "vxe:{}").replace("vxe:ue.vxe", "vxe:{}")
    tmp = ROOT / "_tmp_i18n_obj.cjs"
    out = ROOT / "_tmp_i18n_parsed.json"
    tmp.write_text(
        "const fs = require('fs');\n"
        "const path = require('path');\n"
        f"const o = {fixed};\n"
        "fs.writeFileSync(path.join(__dirname, '_tmp_i18n_parsed.json'), JSON.stringify(o), 'utf8');\n",
        encoding="utf-8",
    )
    proc = subprocess.run(
        ["node", str(tmp)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=str(ROOT),
    )
    try:
        tmp.unlink(missing_ok=True)
    except OSError:
        pass
    if proc.returncode != 0:
        raise SystemExit(proc.stderr or proc.stdout or "node failed")
    data = out.read_text(encoding="utf-8")
    out.unlink(missing_ok=True)
    return json.loads(data)


def flatten_strings(obj, prefix: str = "") -> dict[str, str]:
    out: dict[str, str] = {}
    if not isinstance(obj, dict):
        return out
    for k, v in obj.items():
        p = f"{prefix}.{k}" if prefix else k
        if isinstance(v, dict):
            out.update(flatten_strings(v, p))
        elif isinstance(v, str):
            out[p] = v
    return out


def load_standard_xlsx(path: Path) -> dict[str, str]:
    import openpyxl  # type: ignore

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return {}
    header = [str(c).strip() if c is not None else "" for c in rows[0]]
    # Heuristic: first column 中文, second 英文 (or header names)
    zh_col = 0
    en_col = 1
    for i, h in enumerate(header):
        if h in ("中文", "Chinese", "简体", "词条"):
            zh_col = i
        if h in ("英文", "English", "翻译", "译文"):
            en_col = i
    d: dict[str, str] = {}
    for r in rows[1:]:
        if not r or r[zh_col] is None:
            continue
        zh = str(r[zh_col]).strip()
        if not zh:
            continue
        en = "" if en_col >= len(r) or r[en_col] is None else str(r[en_col]).strip()
        d[zh] = en
    return d


def _parse_cli() -> tuple[Path | None, str | None, str | None]:
    """可选: --config path.json  或  extract_compare.py [js] [standard.xlsx]"""
    config_path = None
    args: list[str] = []
    i = 1
    while i < len(sys.argv):
        if sys.argv[i] in ("-c", "--config") and i + 1 < len(sys.argv):
            config_path = sys.argv[i + 1]
            i += 2
        else:
            args.append(sys.argv[i])
            i += 1
    js_arg = args[0] if args and not args[0].lower().endswith(".xlsx") else None
    xlsx_arg = None
    if len(args) >= 2:
        xlsx_arg = args[1]
    elif len(args) == 1 and args[0].lower().endswith(".xlsx"):
        xlsx_arg = args[0]
    return (Path(config_path) if config_path else None, js_arg, xlsx_arg)


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass

    cfg_path, js_cli, xlsx_cli = _parse_cli()
    cfg = load_config(cfg_path)
    paths = get_paths(cfg)

    js_path = paths["js_file"]
    if js_cli:
        js_path = resolve_js_override(js_cli, cfg)
        if not js_path.is_file():
            raise SystemExit(f"JS 文件不存在: {js_path}")

    xlsx_path = paths["standard_xlsx"]
    if xlsx_cli:
        p = Path(xlsx_cli)
        if p.is_absolute():
            xlsx_path = p
        else:
            cand = paths["rule_dir"] / p
            xlsx_path = cand if cand.is_file() else paths["rule_dir"] / p.name
    do_compare = compare_enabled(cfg) and xlsx_path and xlsx_path.is_file()

    out_json = output_json_path(js_path, cfg)
    out_report = output_report_path(js_path, cfg)

    raw = js_path.read_text(encoding="utf-8")
    xe_lit = extract_balanced_object(raw, "const Xe=")
    # 打包后中文包常为 `...,Ke={global:` 而非 `const Ke=`
    ke_start = raw.find("Ke={", raw.find("const Xe="))
    if ke_start == -1:
        raise SystemExit("Ke={ not found after const Xe=")
    ke_lit = extract_balanced_object(raw[ke_start:], "Ke=")
    Xe = parse_js_object_literal(xe_lit)
    Ke = parse_js_object_literal(ke_lit)
    flat_en = flatten_strings(Xe)
    flat_zh = flatten_strings(Ke)
    paths = sorted(set(flat_en) | set(flat_zh))
    pairs_by_path = []
    for p in paths:
        zh = flat_zh.get(p)
        en = flat_en.get(p)
        if zh is not None and en is not None:
            pairs_by_path.append({"path": p, "zh": zh, "en": en})

    by_zh: dict[str, list[tuple[str, str]]] = {}
    for item in pairs_by_path:
        by_zh.setdefault(item["zh"], []).append((item["en"], item["path"]))
    zh_conflicts = []
    zh_to_en: dict[str, str] = {}
    for zh, arr in by_zh.items():
        ens = sorted({e for e, _ in arr})
        if len(ens) > 1:
            zh_conflicts.append({"zh": zh, "ens": ens, "paths": [x[1] for x in arr]})
        zh_to_en[zh] = arr[-1][0]

    payload = {
        "meta": {
            "source": str(js_path),
            "pairCountByPath": len(pairs_by_path),
            "uniqueZhCount": len(zh_to_en),
            "zhWithMultipleEn": len(zh_conflicts),
        },
        "pairsByPath": pairs_by_path,
        "zhToEn": zh_to_en,
        "zhConflicts": zh_conflicts,
    }
    out_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    lines: list[str] = []
    lines.append("=== JS 提取统计 ===")
    lines.append(json.dumps(payload["meta"], ensure_ascii=False, indent=2))
    if zh_conflicts:
        lines.append("\n=== 同一中文对应多个英文（JS 内路径不同）===")
        for c in zh_conflicts:
            lines.append(f"- 中文: {c['zh']}")
            lines.append(f"  英文候选: {c['ens']}")

    std: dict[str, str] = {}
    if do_compare:
        std = load_standard_xlsx(xlsx_path)
        lines.append(f"\n=== 标准表（Excel）词条数: {len(std)} ===")
        lines.append(f"文件: {xlsx_path}")

        js_zh = set(zh_to_en)
        std_zh = set(std)
        both = js_zh & std_zh
        only_js = js_zh - std_zh
        only_std = std_zh - js_zh

        identical = []
        mismatch = []
        for z in sorted(both):
            a, b = zh_to_en[z], std[z]
            if a == b:
                identical.append((z, a))
            else:
                mismatch.append((z, a, b))

        lines.append("\n--- 完全一致（中文在两份中且英文相同）---")
        lines.append(f"条数: {len(identical)}")
        for z, en in identical:
            lines.append(f"{z}\t{en}")

        lines.append("\n--- 英文不一致（需按标准修正）---")
        lines.append(f"条数: {len(mismatch)}")
        for z, js_en, std_en in mismatch:
            lines.append(f"中文: {z}")
            lines.append(f"  JS:   {js_en}")
            lines.append(f"  标准: {std_en}")

        lines.append("\n--- JS 有、标准无 ---")
        lines.append(f"条数: {len(only_js)}")
        for z in sorted(only_js):
            lines.append(f"{z}\t{zh_to_en[z]}")

        lines.append("\n--- 标准有、JS 无 ---")
        lines.append(f"条数: {len(only_std)}")
        for z in sorted(only_std):
            lines.append(f"{z}\t{std[z]}")
    else:
        if xlsx_path and not xlsx_path.is_file():
            lines.append(f"\n未找到 Excel 标准文件: {xlsx_path}")
        elif not compare_enabled(cfg):
            lines.append("\nconfig.compare.enabled 为 false，已跳过与标准表对比。")
        else:
            lines.append("\n未配置 paths.standard_xlsx，已跳过对比。")
        lines.append("请在 config.json 中设置 paths.standard_xlsx 后重新运行。")

    summary = (
        f"JS 按路径中英配对 {payload['meta']['pairCountByPath']} 条，去重中文 {payload['meta']['uniqueZhCount']} 条。"
    )
    if do_compare:
        summary += (
            f" 与标准表 {len(std)} 条对比：完全一致 {len(identical)}，英文不一致 {len(mismatch)}，"
            f"仅 JS 有 {len(only_js)}，仅标准有 {len(only_std)}。"
        )
    else:
        summary += " 未读取标准 Excel，无法输出差异对比。"
    lines.append("\n=== 一句话总结 ===")
    lines.append(summary)

    out_report.write_text("\n".join(lines), encoding="utf-8")
    print(summary)
    print(f"Wrote: {out_json}")
    print(f"Wrote: {out_report}")


if __name__ == "__main__":
    main()
