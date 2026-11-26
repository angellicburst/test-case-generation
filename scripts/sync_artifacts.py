#!/usr/bin/env python3
"""Sync an edited Markdown document back to Markdown and Excel atomically."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

from case_baseline import (
    assign_missing_ids,
    canonical_tree_order,
    collect_structure_warnings,
    merge_artifacts_to_ir,
    renumber_cases,
    validate_baseline,
)
from create_excel_cases import build_workbook_from_ir, save_workbook_atomic
from render_markdown_cases import render_markdown
from traceability_matrix import sync_traceability_from_cases
from validate_traceability_bundle import summarize, validate_bundle
from validation_utils import (
    ALLOWED_DESIGN_METHODS,
    commit_files_atomic,
    configure_utf8_stdio,
    sibling_temp_path,
    write_text_atomic,
)

GATE_EXITS = (
    "请选择其一：\n"
    "1. 追加 --default-design-method 指定默认设计方法（如 等价类划分）；\n"
    "2. 去掉 --excel，仅同步 Markdown；\n"
    "3. 先用 case_baseline.py export 导出 IR 补全 design_method，再用 render_cases.py 渲染。"
)


def prepare_sync_ir(
    markdown_path: Path,
    excel_path: Path | None = None,
    assign_ids: bool = False,
    renumber: bool = False,
    id_prefix: str | None = None,
    default_design_method: str | None = None,
    sheet_name: str = "测试用例",
) -> tuple[dict[str, Any], list[str], list[str]]:
    """Merge the artifacts and settle IDs, traceability links and design methods."""
    excel_exists = excel_path is not None and excel_path.is_file()
    ir, errors, warnings = merge_artifacts_to_ir(
        md_path=markdown_path,
        excel_path=excel_path if excel_exists else None,
        primary="markdown",
        sheet_name=sheet_name,
        allow_unassigned_ids=assign_ids,
        strict_count=False,
    )
    if errors:
        return {}, errors, warnings
    if not excel_exists:
        ir["meta"]["coverage"] = ""
        ir["meta"]["notes"] = ""

    if assign_ids:
        assign_missing_ids(ir["cases"], id_prefix)
    ir["cases"] = canonical_tree_order(ir["cases"])
    warnings.extend(collect_structure_warnings(ir["cases"]))
    if renumber:
        renumber_cases(ir["cases"], ir["traceability"], id_prefix)

    entries, sync_errors, sync_warnings = sync_traceability_from_cases(
        ir["traceability"], ir["cases"]
    )
    warnings.extend(sync_warnings)
    if sync_errors:
        return {}, sync_errors, warnings
    ir["traceability"] = list(entries.values())

    lost = [entry["id"] for entry in ir["traceability"] if entry.get("status") == "待覆盖"]
    if lost:
        return {}, [
            "以下追溯条目失去全部关联用例，已阻断同步: "
            + "、".join(lost)
            + "。请补充用例或删除该条目。"
        ], warnings

    if excel_path is not None:
        if default_design_method is not None and default_design_method not in ALLOWED_DESIGN_METHODS:
            return {}, [
                f"--default-design-method 必须为 {'、'.join(ALLOWED_DESIGN_METHODS)} 之一，"
                f"实际为“{default_design_method}”。"
            ], warnings
        invalid = [
            case["id"] or "无 ID 用例"
            for case in ir["cases"]
            if case["design_method"] not in ALLOWED_DESIGN_METHODS
        ]
        if invalid:
            if default_design_method is None:
                return {}, [
                    "以下用例缺少合法的 design_method，已拒绝写入: "
                    + "、".join(invalid)
                    + "。\n"
                    + GATE_EXITS
                ], warnings
            for case in ir["cases"]:
                if case["design_method"] not in ALLOWED_DESIGN_METHODS:
                    case["design_method"] = default_design_method
            warnings[:] = [
                warning
                for warning in warnings
                if not warning.startswith("以下用例缺少 design_method")
            ]

    errors = validate_baseline(ir, require_design_method=excel_path is not None)
    if errors:
        return {}, errors, warnings
    return ir, [], warnings


def main(argv: list[str] | None = None) -> int:
    configure_utf8_stdio()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--from-markdown", required=True, help="人工编辑后的 Markdown 文件")
    parser.add_argument("--excel", help="既有或待生成的 Excel 文件")
    parser.add_argument("--assign-ids", action="store_true", help="为新增用例补号")
    parser.add_argument("--renumber", action="store_true", help="定稿前重排全部用例 ID")
    parser.add_argument("--id-prefix", help="补号或重排使用的需求简写前缀")
    parser.add_argument("--default-design-method", help="回填缺失设计方法时使用的默认值")
    parser.add_argument("--sheet-name", default="测试用例", help="主工作表名称")
    args = parser.parse_args(argv)

    markdown_path = Path(args.from_markdown).expanduser().resolve()
    excel_path = Path(args.excel).expanduser().resolve() if args.excel else None
    if not markdown_path.is_file():
        print(f"Markdown 文件不存在: {markdown_path}", file=sys.stderr)
        return 1
    if excel_path is not None and excel_path.exists() and excel_path.is_dir():
        print(f"Excel 路径必须是文件路径: {excel_path}", file=sys.stderr)
        return 1

    ir, errors, warnings = prepare_sync_ir(
        markdown_path,
        excel_path=excel_path,
        assign_ids=args.assign_ids,
        renumber=args.renumber,
        id_prefix=args.id_prefix,
        default_design_method=args.default_design_method,
        sheet_name=args.sheet_name,
    )
    if errors:
        print("同步失败，未写出任何文件:", file=sys.stderr)
        for error in errors:
            print(f"- {error}", file=sys.stderr)
        return 1
    if excel_path is None or not excel_path.is_file():
        ir["meta"]["sheet_name"] = args.sheet_name

    staged: list[tuple[Path, Path]] = []
    staged_markdown: Path | None = None
    staged_excel: Path | None = None
    sheet_name = ir["meta"]["sheet_name"]
    try:
        markdown_path.parent.mkdir(parents=True, exist_ok=True)
        if excel_path is not None:
            excel_path.parent.mkdir(parents=True, exist_ok=True)
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
        print(f"同步失败，未写出任何文件: {exc}", file=sys.stderr)
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
        print("同步结果未通过校验，未写出任何文件:", file=sys.stderr)
        for error in bundle_errors:
            print(f"- {error}", file=sys.stderr)
        return 1

    commit_errors = commit_files_atomic(staged)
    if commit_errors:
        print("写出文件失败:", file=sys.stderr)
        for error in commit_errors:
            print(f"- {error}", file=sys.stderr)
        rollback_failed = any("回滚" in error and "失败" in error for error in commit_errors)
        if rollback_failed:
            print(
                "提示：部分文件回滚失败，请先按上方备份路径手工恢复。",
                file=sys.stderr,
            )
        else:
            print(
                "提示：其余目标文件均未变更（已保持原样或已回滚）；"
                "请关闭占用程序或修复权限后原样重跑本命令。",
                file=sys.stderr,
            )
        return 1

    for warning in warnings:
        print(f"提示：{warning}", file=sys.stderr)
    cases, covered, pending = summarize(ir)
    outputs = "、".join(str(path) for path in (markdown_path, excel_path) if path)
    print(
        f"同步完成 {outputs}：共 {cases} 条用例，"
        f"已覆盖追溯 {covered} 条，待确认 {pending} 条。"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
