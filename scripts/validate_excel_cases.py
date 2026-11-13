#!/usr/bin/env python3
"""Validate structure, headers, mandatory fields, and uniqueness for Excel (.xlsx) test cases."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

EXPECTED_HEADERS = [
    "需求/追踪 ID",
    "用例 ID",
    "用例标题",
    "用例类型",
    "设计方法",
    "优先级",
    "前置条件",
    "测试数据",
    "执行步骤",
    "预期结果",
    "备注",
]

MANDATORY_FIELDS = [
    "需求/追踪 ID",
    "用例 ID",
    "用例标题",
    "用例类型",
    "设计方法",
    "优先级",
    "前置条件",
    "执行步骤",
    "预期结果",
]

ALLOWED_PRIORITIES = {"P0", "P1", "P2", "P3", "P4"}


def validate_excel_file(file_path: Path, sheet_name: str = "测试用例") -> list[str]:
    errors: list[str] = []

    if not file_path.is_file():
        return [f"文件不存在: {file_path}"]

    try:
        from openpyxl import load_workbook
    except ImportError:
        return ["缺少 openpyxl 依赖，无法读取和校验 .xlsx 文件，请执行: pip install openpyxl"]

    try:
        wb = load_workbook(filename=str(file_path), data_only=True)
    except Exception as exc:
        return [f"无法读取 Excel 文件: {exc}"]

    if sheet_name not in wb.sheetnames:
        return [f"工作簿中缺少名为 '{sheet_name}' 的工作表，当前包含的工作表: {wb.sheetnames}"]

    if "生成说明" not in wb.sheetnames:
        errors.append(
            f"工作簿中缺少名为 '生成说明' 的工作表，当前包含的工作表: {wb.sheetnames}"
        )

    ws = wb[sheet_name]

    # 检查第 2 行为列头
    row2_values = [ws.cell(row=2, column=col).value for col in range(1, len(EXPECTED_HEADERS) + 1)]
    row2_cleaned = [str(v).strip() if v is not None else "" for v in row2_values]

    if row2_cleaned != EXPECTED_HEADERS:
        errors.append(
            f"第 2 行表头不符合规范。\n期望: {EXPECTED_HEADERS}\n实际: {row2_cleaned}"
        )

    # 列索引映射
    header_to_col = {header: idx + 1 for idx, header in enumerate(EXPECTED_HEADERS)}

    seen_case_ids: set[str] = set()
    seen_titles: set[str] = set()
    valid_rows_count = 0

    max_row = ws.max_row or 0
    if max_row < 3:
        errors.append("工作表中未发现有效用例数据行（数据应从第 3 行开始）。")
        return errors

    for row_idx in range(3, max_row + 1):
        # 检查该行是否全空
        row_vals = [ws.cell(row=row_idx, column=c).value for c in range(1, len(EXPECTED_HEADERS) + 1)]
        if not any(v is not None and str(v).strip() != "" for v in row_vals):
            continue  # 跳过全空行

        valid_rows_count += 1

        # 检查必填项
        for field in MANDATORY_FIELDS:
            col_idx = header_to_col[field]
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is None or not str(val).strip():
                errors.append(f"第 {row_idx} 行缺少必填字段: '{field}'。")

        # 检查用例 ID
        case_id_val = ws.cell(row=row_idx, column=header_to_col["用例 ID"]).value
        if case_id_val is not None and str(case_id_val).strip():
            cid = str(case_id_val).strip()
            if cid in seen_case_ids:
                errors.append(f"第 {row_idx} 行用例 ID 重复: '{cid}'。")
            seen_case_ids.add(cid)

        # 检查用例标题
        title_val = ws.cell(row=row_idx, column=header_to_col["用例标题"]).value
        if title_val is not None and str(title_val).strip():
            title_str = str(title_val).strip()
            if title_str in seen_titles:
                errors.append(f"第 {row_idx} 行用例标题重复: '{title_str}'。")
            seen_titles.add(title_str)

        # 检查优先级
        priority_val = ws.cell(row=row_idx, column=header_to_col["优先级"]).value
        if priority_val is not None and str(priority_val).strip():
            p_str = str(priority_val).strip().upper()
            if p_str not in ALLOWED_PRIORITIES:
                errors.append(
                    f"第 {row_idx} 行优先级无效: '{priority_val}'，必须为 {ALLOWED_PRIORITIES} 之一。"
                )

    if valid_rows_count == 0:
        errors.append("工作表中未包含任何有效用例数据行。")

    return errors


def main() -> None:
    parser = argparse.ArgumentParser(description="校验 Excel (.xlsx) 测试用例文件的结构、表头与必填字段。")
    parser.add_argument("path", nargs="?", help="要校验的 .xlsx 文件路径")
    parser.add_argument("--path", dest="opt_path", help="要校验的 .xlsx 文件路径（兼容选项）")
    parser.add_argument("--sheet-name", default="测试用例", help="测试用例工作表名称（默认为'测试用例'）")
    args = parser.parse_args()

    target_path_str = args.opt_path or args.path
    if not target_path_str:
        parser.error("必须提供 .xlsx 文件路径")

    target_path = Path(target_path_str).expanduser().resolve()
    errors = validate_excel_file(target_path, sheet_name=args.sheet_name)

    if errors:
        print(f"校验失败: 发现 {len(errors)} 处错误:\n", file=sys.stderr)
        for idx, err in enumerate(errors, start=1):
            print(f"{idx}. {err}", file=sys.stderr)
        sys.exit(1)

    print(f"校验通过: {target_path}")
    sys.exit(0)


if __name__ == "__main__":
    main()

