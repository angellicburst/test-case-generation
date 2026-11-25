#!/usr/bin/env python3
"""Validate Markdown and/or Excel test-case artifacts and their consistency."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

from case_baseline import (
    describe_ir_differences,
    load_baseline_raw,
    normalize_baseline,
    parse_excel_to_ir,
    parse_markdown_to_ir,
    validate_baseline,
)
from validate_excel_cases import validate_excel_file
from validate_markdown_lists import validate as validate_markdown
from validation_utils import configure_utf8_stdio


def _read_markdown(path: Path) -> tuple[str, list[str]]:
    try:
        return path.read_text(encoding="utf-8-sig"), []
    except (OSError, UnicodeDecodeError) as exc:
        return "", [f"无法读取 Markdown 文件: {path}（{exc}）"]


def summarize(ir: dict[str, Any]) -> tuple[int, int, int]:
    entries = [entry for entry in ir.get("traceability") or [] if isinstance(entry, dict)]
    covered = sum(1 for entry in entries if entry.get("status") == "已覆盖")
    pending = sum(1 for entry in entries if entry.get("status") == "待确认")
    return len(ir.get("cases") or []), covered, pending


def validate_bundle(
    markdown_path: Path | None = None,
    excel_path: Path | None = None,
    baseline_path: Path | None = None,
    sheet_name: str = "测试用例",
    coverage_only: bool = False,
) -> list[str]:
    """Run self-contained validation and, with both files, cross-format checks."""
    if markdown_path is None and excel_path is None:
        return ["必须提供 Markdown 或 Excel 至少一个文件。"]

    errors: list[str] = []
    baseline: dict[str, Any] | None = None
    if baseline_path is not None:
        try:
            raw_baseline = load_baseline_raw(baseline_path)
        except ValueError as exc:
            return [str(exc)]
        errors.extend(validate_baseline(raw_baseline))
        if errors:
            return errors
        baseline = normalize_baseline(raw_baseline)

    if not coverage_only:
        if markdown_path is not None:
            text, read_errors = _read_markdown(markdown_path)
            errors.extend(read_errors)
            if not read_errors:
                errors.extend(validate_markdown(text))
        if excel_path is not None:
            errors.extend(validate_excel_file(excel_path, sheet_name=sheet_name))

    markdown_ir: dict[str, Any] = {}
    excel_ir: dict[str, Any] = {}
    if markdown_path is not None:
        markdown_ir, markdown_errors, _ = parse_markdown_to_ir(markdown_path)
        if coverage_only:
            errors.extend(markdown_errors)
    if excel_path is not None:
        excel_ir, excel_errors, _ = parse_excel_to_ir(excel_path, sheet_name=sheet_name)
        if coverage_only:
            errors.extend(excel_errors)

    if baseline is not None:
        if markdown_ir:
            errors.extend(
                describe_ir_differences(
                    baseline,
                    markdown_ir,
                    "基线 IR",
                    "Markdown",
                    include_design_method=False,
                )
            )
        if excel_ir:
            errors.extend(
                describe_ir_differences(
                    baseline,
                    excel_ir,
                    "基线 IR",
                    "Excel",
                    include_tree_fields=False,
                )
            )

    if markdown_ir and excel_ir:
        errors.extend(
            describe_ir_differences(
                markdown_ir,
                excel_ir,
                "Markdown",
                "Excel",
                include_tree_fields=False,
                include_design_method=False,
            )
        )

    return list(dict.fromkeys(errors))


def main(argv: list[str] | None = None) -> int:
    configure_utf8_stdio()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--markdown", type=Path, help="Markdown 测试用例文件")
    parser.add_argument("--excel", type=Path, help="Excel 测试用例文件")
    parser.add_argument("--baseline", type=Path, help="可选的合并 IR，用于额外自校验")
    parser.add_argument("--sheet-name", default="测试用例", help="Excel 主工作表名")
    parser.add_argument(
        "--coverage-only",
        action="store_true",
        help="跳过格式细节，仅核对两种格式的用例与追溯是否一致",
    )
    args = parser.parse_args(argv)

    if args.markdown is None and args.excel is None:
        parser.error("至少提供 --markdown 或 --excel 之一。")
    if args.coverage_only and (args.markdown is None or args.excel is None):
        parser.error("--coverage-only 必须同时提供 --markdown 和 --excel。")

    errors = validate_bundle(
        markdown_path=args.markdown,
        excel_path=args.excel,
        baseline_path=args.baseline,
        sheet_name=args.sheet_name,
        coverage_only=args.coverage_only,
    )
    if errors:
        print(f"追溯总校验失败: 发现 {len(errors)} 处错误:", file=sys.stderr)
        for index, error in enumerate(errors, start=1):
            print(f"{index}. {error}", file=sys.stderr)
        return 1

    ir: dict[str, Any] = {}
    if args.markdown is not None:
        ir, _, _ = parse_markdown_to_ir(args.markdown)
    elif args.excel is not None:
        ir, _, _ = parse_excel_to_ir(args.excel, sheet_name=args.sheet_name)
    cases, covered, pending = summarize(ir)
    print(f"校验通过：共 {cases} 条用例，已覆盖追溯 {covered} 条，待确认 {pending} 条。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
