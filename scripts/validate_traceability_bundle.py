#!/usr/bin/env python3
"""Validate traceability completeness across Markdown and Excel test-case files."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from traceability_matrix import final_matrix_errors, load_traceability_matrix
from validate_excel_cases import validate_excel_file
from validate_markdown_lists import validate as validate_markdown


def validate_coverage(matrix: dict[str, dict[str, object]]) -> list[str]:
    """Validate that every covered requirement is linked by both formats."""
    errors = final_matrix_errors(matrix)
    covered_ids = {
        entry["id"]
        for entry in matrix.values()
        if entry["status"] == "已覆盖"
    }
    markdown_ids = {
        entry["id"]
        for entry in matrix.values()
        if entry["case_titles"]
    }
    excel_ids = {
        entry["id"]
        for entry in matrix.values()
        if entry["case_ids"]
    }
    for trace_id in sorted(covered_ids - markdown_ids):
        errors.append(f"已覆盖追溯项“{trace_id}”缺少 Markdown 用例关联。")
    for trace_id in sorted(covered_ids - excel_ids):
        errors.append(f"已覆盖追溯项“{trace_id}”缺少 Excel 用例关联。")
    for trace_id in sorted(markdown_ids ^ excel_ids):
        errors.append(f"追溯项“{trace_id}”在 Markdown 与 Excel 的覆盖不一致。")
    return errors


def validate_bundle(
    matrix_path: Path,
    markdown_path: Path | None = None,
    excel_path: Path | None = None,
    sheet_name: str = "测试用例",
    coverage_only: bool = False,
) -> list[str]:
    matrix, errors = load_traceability_matrix(matrix_path)
    if errors:
        return errors
    if coverage_only and (markdown_path is None or excel_path is None):
        return ["仅覆盖核对必须同时提供 Markdown 和 Excel 文件路径。"]
    if coverage_only:
        for label, path in (("Markdown", markdown_path), ("Excel", excel_path)):
            if not path.is_file():
                errors.append(f"{label} 文件不存在: {path}")
        if errors:
            return errors

    if not coverage_only and markdown_path is not None:
        if not markdown_path.is_file():
            errors.append(f"Markdown 文件不存在: {markdown_path}")
        else:
            try:
                markdown_text = markdown_path.read_text(encoding="utf-8-sig")
            except UnicodeDecodeError as exc:
                errors.append(f"Markdown 文件编码不是 UTF-8: {markdown_path}（{exc}）")
            else:
                errors.extend(validate_markdown(markdown_text, matrix))

    if not coverage_only and excel_path is not None:
        errors.extend(validate_excel_file(excel_path, matrix, sheet_name=sheet_name))

    if markdown_path is not None and excel_path is not None:
        errors.extend(validate_coverage(matrix))

    return list(dict.fromkeys(errors))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--traceability-matrix", type=Path, required=True)
    parser.add_argument("--markdown", type=Path)
    parser.add_argument("--excel", type=Path)
    parser.add_argument("--sheet-name", default="测试用例")
    parser.add_argument(
        "--coverage-only",
        action="store_true",
        help="仅核对矩阵的 Markdown/Excel 覆盖关联，依赖两种格式已分别校验通过。",
    )
    args = parser.parse_args()

    if args.markdown is None and args.excel is None:
        parser.error("至少提供 --markdown 或 --excel 之一。")
    if args.coverage_only and (args.markdown is None or args.excel is None):
        parser.error("--coverage-only 必须同时提供 --markdown 和 --excel。")

    errors = validate_bundle(
        args.traceability_matrix,
        markdown_path=args.markdown,
        excel_path=args.excel,
        sheet_name=args.sheet_name,
        coverage_only=args.coverage_only,
    )
    if errors:
        print(f"追溯总校验失败: 发现 {len(errors)} 处错误:\n", file=sys.stderr)
        for index, error in enumerate(errors, start=1):
            print(f"{index}. {error}", file=sys.stderr)
        return 1

    print("追溯总校验通过。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
