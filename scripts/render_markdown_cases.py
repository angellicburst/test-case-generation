#!/usr/bin/env python3
"""Render the merged v2 IR into the standard Markdown task tree."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

from case_baseline import (
    canonical_tree_order,
    load_baseline_raw,
    prepare_render_ir,
)
from validation_utils import (
    configure_utf8_stdio,
    escape_md_table_cell,
    normalize_list_items,
    write_text_atomic,
)

COVERED_HEADERS = ("追溯 ID", "来源", "定位", "规则", "关联用例")
PENDING_HEADERS = ("追溯 ID", "来源", "定位", "未明确项", "当前测试假设", "确认问题")
FIELDS = (
    ("前置操作", "pre_operations", True),
    ("测试数据", "test_data", False),
    ("测试步骤", "steps", True),
    ("预期结果", "expected_results", True),
    ("后置操作", "post_operations", False),
)


def _render_items(items: list[str], indent: int) -> list[str]:
    lines: list[str] = []
    prefix = " " * indent
    if len(items) == 1:
        lines.append(f"{prefix}- {items[0]}")
        return lines
    for index, item in enumerate(items, start=1):
        lines.append(f"{prefix}{index}. {item}")
    return lines


def _render_table(headers: tuple[str, ...], rows: list[list[str]]) -> list[str]:
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join(":---" for _ in headers) + " |",
    ]
    for row in rows:
        lines.append("| " + " | ".join(escape_md_table_cell(cell) for cell in row) + " |")
    return lines


def render_markdown(ir: dict[str, Any]) -> str:
    meta = ir["meta"]
    cases = ir["cases"]
    trace_order = {entry["id"]: index for index, entry in enumerate(ir["traceability"])}
    default_order = len(trace_order)

    lines: list[str] = [f"# {meta['title']}", "", f"> 用例数量: 用例 {len(cases)} 条", ""]

    tree: dict[str, dict[str, dict[str, list[dict[str, Any]]]]] = {}
    for case in canonical_tree_order(cases):
        tree.setdefault(case["module"], {}).setdefault(case["type"], {}).setdefault(
            case["group"], []
        ).append(case)

    for module, types in tree.items():
        lines.extend([f"- [ ] {module}", ""])
        for type_name, groups in types.items():
            lines.extend([f"  - [ ] {type_name}", ""])
            for group, group_cases in groups.items():
                if group:
                    lines.extend([f"    - [ ] {group}", ""])
                case_indent = 6 if group else 4
                field_indent = case_indent + 2
                for case in group_cases:
                    lines.extend(
                        [
                            " " * case_indent
                            + f"- [ ] {case['polarity']}-{case['title']}-{case['priority']}",
                            "",
                        ]
                    )
                    refs = sorted(
                        case["traceability_refs"],
                        key=lambda ref: trace_order.get(ref, default_order),
                    )
                    lines.extend(
                        [
                            " " * field_indent
                            + f"> [{case['id']}] 追溯: {', '.join(refs)}",
                            "",
                        ]
                    )
                    for label, key, required in FIELDS:
                        items = normalize_list_items(case.get(key))
                        if not items:
                            if required:
                                lines.extend([" " * field_indent + f"- {label}", ""])
                            continue
                        lines.extend([" " * field_indent + f"- {label}", ""])
                        lines.extend(_render_items(items, case_indent + 4))
                        lines.append("")

    lines.extend(["## 需求追溯", "", "### 需求追溯表", ""])
    covered_rows = [
        [
            entry["id"],
            entry["source"],
            entry["location"],
            entry["rule"],
            ", ".join(entry.get("case_ids", [])),
        ]
        for entry in ir["traceability"]
        if entry.get("status") != "待确认"
    ]
    lines.extend(_render_table(COVERED_HEADERS, covered_rows))
    lines.append("")

    pending = [entry for entry in ir["traceability"] if entry.get("status") == "待确认"]
    if pending:
        lines.extend(["### 待确认项", ""])
        pending_rows = [
            [
                entry["id"],
                entry["source"],
                entry["location"],
                entry.get("issue", ""),
                entry.get("assumption", ""),
                entry.get("question", ""),
            ]
            for entry in pending
        ]
        lines.extend(_render_table(PENDING_HEADERS, pending_rows))
        lines.append("")

    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines) + "\n"


def main(argv: list[str] | None = None) -> int:
    configure_utf8_stdio()
    parser = argparse.ArgumentParser(description="将合并 IR 渲染为 Markdown 测试用例。")
    parser.add_argument("--baseline", required=True, help="合并 IR JSON 路径")
    parser.add_argument("--output", required=True, help="输出 Markdown 文件路径")
    parser.add_argument("--force", action="store_true", help="输出文件已存在时确认覆盖")
    args = parser.parse_args(argv)

    try:
        raw = load_baseline_raw(args.baseline)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    ir, errors, warnings = prepare_render_ir(raw, require_design_method=False)
    if errors:
        print("\n".join(errors), file=sys.stderr)
        return 1

    output = Path(args.output).expanduser().resolve()
    if output.exists() and output.is_dir():
        print("输出路径必须是文件路径，不能是目录。", file=sys.stderr)
        return 1
    if output.exists() and not args.force:
        print(f"输出文件已存在: {output}。确认覆盖后请追加 --force。", file=sys.stderr)
        return 1

    for warning in warnings:
        print(f"提示：{warning}", file=sys.stderr)

    try:
        output.parent.mkdir(parents=True, exist_ok=True)
        write_text_atomic(output, render_markdown(ir))
    except (OSError, RuntimeError) as exc:
        print(str(exc), file=sys.stderr)
        return 1
    print(f"成功生成 Markdown 测试用例: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
