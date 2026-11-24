#!/usr/bin/env python3
"""Render a merged v2 IR into Markdown and/or Excel and validate the result."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

from case_baseline import (
    assign_missing_ids,
    canonical_tree_order,
    load_baseline_raw,
    normalize_baseline,
    validate_baseline,
)
from create_excel_cases import build_workbook_from_ir, save_workbook_atomic
from render_markdown_cases import render_markdown
from traceability_matrix import sync_traceability_from_cases
from validate_traceability_bundle import summarize, validate_bundle
from validation_utils import (
    commit_files_atomic,
    configure_utf8_stdio,
    sibling_temp_path,
    write_text_atomic,
)


def build_render_ir(
    raw: Any,
    assign_ids: bool = False,
    id_prefix: str | None = None,
) -> tuple[dict[str, Any], list[str], list[str]]:
    """Normalize the IR, settle IDs and traceability links before rendering."""
    errors = validate_baseline(
        raw,
        allow_unassigned_ids=assign_ids,
        require_finalized_traceability=False,
        check_traceability_links=False,
    )
    if errors:
        return {}, errors, []

    ir = normalize_baseline(raw)
    if assign_ids:
        assign_missing_ids(ir["cases"], id_prefix)
    ir["cases"] = canonical_tree_order(ir["cases"])

    entries, sync_errors, warnings = sync_traceability_from_cases(
        ir["traceability"], ir["cases"]
    )
    if sync_errors:
        return {}, sync_errors, warnings
    ir["traceability"] = list(entries.values())
    return ir, [], warnings


def main(argv: list[str] | None = None) -> int:
    configure_utf8_stdio()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline", required=True, help="合并 IR JSON 文件路径")
    parser.add_argument("--markdown", help="输出的 Markdown 文件路径")
    parser.add_argument("--excel", help="输出的 Excel 文件路径")
    parser.add_argument("--assign-ids", action="store_true", help="自动为缺少 ID 的用例补号")
    parser.add_argument("--id-prefix", help="补号使用的需求简写前缀")
    parser.add_argument(
        "--cleanup-baseline",
        action="store_true",
        help="全部校验通过后删除临时 IR 文件",
    )
    parser.add_argument("--force", action="store_true", help="输出文件已存在时确认覆盖")
    args = parser.parse_args(argv)

    if args.markdown is None and args.excel is None:
        parser.error("--markdown 与 --excel 至少指定一个。")

    baseline_path = Path(args.baseline).expanduser().resolve()
    markdown_path = Path(args.markdown).expanduser().resolve() if args.markdown else None
    excel_path = Path(args.excel).expanduser().resolve() if args.excel else None

    for target in (markdown_path, excel_path):
        if target is None:
            continue
        if target.exists() and target.is_dir():
            print(f"输出路径必须是文件路径，不能是目录: {target}", file=sys.stderr)
            return 1
        if target.exists() and not args.force:
            print(f"输出文件已存在: {target}。确认覆盖后请追加 --force。", file=sys.stderr)
            return 1

    try:
        raw = load_baseline_raw(baseline_path)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    ir, errors, warnings = build_render_ir(raw, args.assign_ids, args.id_prefix)
    if not errors:
        errors = validate_baseline(ir, require_design_method=excel_path is not None)
    if errors:
        print("渲染失败，未写出任何文件:", file=sys.stderr)
        for error in errors:
            print(f"- {error}", file=sys.stderr)
        return 1

    staged: list[tuple[Path, Path]] = []
    staged_markdown: Path | None = None
    staged_excel: Path | None = None
    sheet_name = ir["meta"]["sheet_name"]
    try:
        for target in (markdown_path, excel_path):
            if target is not None:
                target.parent.mkdir(parents=True, exist_ok=True)
        if markdown_path is not None:
            temp = sibling_temp_path(markdown_path)
            write_text_atomic(temp, render_markdown(ir))
            staged.append((markdown_path, temp))
            staged_markdown = temp
        if excel_path is not None:
            temp = sibling_temp_path(excel_path)
            save_workbook_atomic(build_workbook_from_ir(ir), temp)
            staged.append((excel_path, temp))
            staged_excel = temp
    except (OSError, RuntimeError, ValueError) as exc:
        for _target, temp in staged:
            try:
                temp.unlink(missing_ok=True)
            except OSError:
                pass
        print(f"渲染失败，未写出任何文件: {exc}", file=sys.stderr)
        return 1

    bundle_errors = validate_bundle(
        markdown_path=staged_markdown,
        excel_path=staged_excel,
        sheet_name=sheet_name,
    )
    if bundle_errors:
        for _target, temp in staged:
            try:
                temp.unlink(missing_ok=True)
            except OSError:
                pass
        print("渲染结果未通过校验，未写出任何文件:", file=sys.stderr)
        for error in bundle_errors:
            print(f"- {error}", file=sys.stderr)
        return 1

    commit_errors = commit_files_atomic(staged)
    if commit_errors:
        print("写出文件失败:", file=sys.stderr)
        for error in commit_errors:
            print(f"- {error}", file=sys.stderr)
        return 1

    if args.cleanup_baseline:
        try:
            baseline_path.unlink(missing_ok=True)
        except OSError as exc:
            print(f"提示：临时 IR 未删除: {baseline_path}（{exc}）", file=sys.stderr)

    for warning in warnings:
        print(f"提示：{warning}", file=sys.stderr)
    cases, covered, pending = summarize(ir)
    outputs = "、".join(str(path) for path in (markdown_path, excel_path) if path)
    print(
        f"已生成 {outputs}：共 {cases} 条用例，"
        f"已覆盖追溯 {covered} 条，待确认 {pending} 条。"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
