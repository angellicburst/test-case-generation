#!/usr/bin/env python3
"""Validate a self-contained Excel (.xlsx) v2 test-case workbook."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Any

from case_baseline import (
    INFO_LABELS,
    RESERVED_SHEET_NAMES,
    V2_HEADERS,
    parse_excel_to_ir,
    validate_baseline,
)
from validation_utils import configure_utf8_stdio, is_valid_date_or_placeholder

REQUIRED_INFO_LABELS = ("需求名称", "上线日期", "用例数量")
COUNT_PATTERN = re.compile(r"[0-9]+")
REQUIREMENTS_FILE = Path(__file__).resolve().parents[1] / "requirements.txt"


def _cell_text(value: Any) -> str:
    return str(value).strip() if value is not None else ""


def _column_letter(index: int) -> str:
    return chr(64 + index)


def _row_has_content(ws: Any, row: int, width: int) -> bool:
    return any(
        ws.cell(row=row, column=column).value is not None
        and str(ws.cell(row=row, column=column).value).strip()
        for column in range(1, width + 1)
    )


def _last_data_row(ws: Any, width: int) -> int:
    rows = [
        row
        for row in range(3, (ws.max_row or 0) + 1)
        if _row_has_content(ws, row, width)
    ]
    return rows[-1] if rows else 2


def _sheet_layout_errors(
    ws: Any, width: int, merge_range: str, errors: list[str], label: str
) -> None:
    merges = {str(merged) for merged in ws.merged_cells.ranges}
    if merges != {merge_range}:
        errors.append(
            f"{label}必须恰好包含 {merge_range} 合并标题区域，不得存在其他合并区域。"
        )
    if ws.freeze_panes != "A3":
        errors.append(f"{label}必须冻结窗格至 A3。")
    expected_filter = f"A2:{_column_letter(width)}{_last_data_row(ws, width)}"
    if ws.auto_filter.ref != expected_filter:
        errors.append(
            f"{label}筛选范围必须为 {expected_filter}，"
            f"实际为 {ws.auto_filter.ref or '未设置'}。"
        )
    max_column = ws.max_column or 0
    if max_column <= width:
        return
    for row in range(1, (ws.max_row or 0) + 1):
        for column in range(width + 1, max_column + 1):
            value = ws.cell(row=row, column=column).value
            if value is not None and str(value).strip():
                errors.append(
                    f"{label}第 {row} 行第 {column} 列存在额外内容，"
                    f"只允许 A–{_column_letter(width)} 共 {width} 列。"
                )
                return


def validate_info_sheet(workbook: Any, expected_count: int) -> list[str]:
    if "生成说明" not in workbook.sheetnames:
        return ["工作簿中缺少名为“生成说明”的工作表。"]

    info = workbook["生成说明"]
    errors: list[str] = []
    label_rows: dict[str, list[int]] = {label: [] for label in INFO_LABELS}
    for row in range(1, (info.max_row or 0) + 1):
        label = _cell_text(info.cell(row=row, column=1).value)
        if label in label_rows:
            label_rows[label].append(row)

    values: dict[str, str] = {}
    for label in INFO_LABELS:
        rows = label_rows[label]
        if not rows:
            errors.append(f"生成说明缺少字段“{label}”。")
        elif len(rows) > 1:
            errors.append(f"生成说明字段“{label}”重复出现 {len(rows)} 次。")
        else:
            values[label] = _cell_text(info.cell(row=rows[0], column=2).value)
            if "[待确认]" in values[label]:
                errors.append(f"生成说明字段“{label}”不得包含 [待确认] 标记。")

    for label in REQUIRED_INFO_LABELS:
        if label in values and not values[label]:
            errors.append(f"生成说明字段“{label}”不能为空。")

    release_date = values.get("上线日期")
    if release_date and not is_valid_date_or_placeholder(release_date):
        errors.append("生成说明字段“上线日期”必须为有效 YYYY-MM-DD 或 未指定日期。")

    count_text = values.get("用例数量")
    if count_text:
        if not COUNT_PATTERN.fullmatch(count_text):
            errors.append(f"生成说明字段“用例数量”必须为整数，实际为“{count_text}”。")
        elif int(count_text) != expected_count:
            errors.append(
                f"生成说明字段“用例数量”为 {count_text}，"
                f"与实际用例行数 {expected_count} 不一致。"
            )

    for row in range(1, (info.max_row or 0) + 1):
        for column in (1, 2):
            if info.cell(row=row, column=column).data_type == "f":
                errors.append(f"生成说明第 {row} 行包含公式，必须使用纯文本。")
    for row in range(1, (info.max_row or 0) + 1):
        for column in range(3, (info.max_column or 0) + 1):
            value = info.cell(row=row, column=column).value
            if value is not None and str(value).strip():
                errors.append(
                    f"生成说明第 {row} 行第 {column} 列存在额外内容，只允许 A–B 两列。"
                )
                return errors
    return errors


def validate_excel_file(file_path: Path, sheet_name: str = "测试用例") -> list[str]:
    """Validate one workbook without any external traceability matrix."""
    errors: list[str] = []
    if not file_path.is_file():
        return [f"文件不存在: {file_path}"]

    try:
        from openpyxl import load_workbook
    except ImportError:
        return [
            "缺少 openpyxl 依赖，无法读取和校验 .xlsx 文件，请执行: "
            f'pip install -r "{REQUIREMENTS_FILE}"'
        ]

    try:
        workbook = load_workbook(filename=str(file_path), data_only=False)
    except Exception as exc:
        return [f"无法读取 Excel 文件: {exc}"]

    if sheet_name in RESERVED_SHEET_NAMES:
        return [f"主工作表名不得使用保留名: {sheet_name}。"]
    if sheet_name not in workbook.sheetnames:
        return [
            f"工作簿中缺少名为“{sheet_name}”的工作表，当前包含: {workbook.sheetnames}"
        ]

    sheet = workbook[sheet_name]
    _sheet_layout_errors(sheet, len(V2_HEADERS), "A1:K1", errors, "主工作表")

    has_pending_sheet = "待确认项" in workbook.sheetnames
    expected_sheets = {sheet_name, "生成说明", "需求追溯", "待确认项"}
    if not has_pending_sheet:
        expected_sheets.discard("待确认项")
    unexpected = set(workbook.sheetnames) - expected_sheets
    if unexpected:
        errors.append(f"工作簿包含多余的工作表: {'、'.join(sorted(unexpected))}。")

    if "需求追溯" in workbook.sheetnames:
        traceability_sheet = workbook["需求追溯"]
        _sheet_layout_errors(
            traceability_sheet, 5, "A1:E1", errors, "“需求追溯”工作表"
        )
        if _last_data_row(traceability_sheet, 5) < 3:
            errors.append("“需求追溯”工作表至少需要 1 行有效数据。")
    if has_pending_sheet:
        pending_sheet = workbook["待确认项"]
        _sheet_layout_errors(pending_sheet, 6, "A1:F1", errors, "“待确认项”工作表")
        if _last_data_row(pending_sheet, 6) < 3:
            errors.append("“待确认项”工作表至少需要 1 行有效数据，禁止空壳工作表。")

    errors.extend(validate_info_sheet(workbook, _count_case_rows(sheet)))

    ir, parse_errors, _warnings = parse_excel_to_ir(
        file_path, sheet_name=sheet_name
    )
    errors.extend(parse_errors)
    if ir:
        errors.extend(
            validate_baseline(
                ir,
                require_design_method=True,
                require_tree_fields=False,
            )
        )
    return list(dict.fromkeys(errors))


def _count_case_rows(ws: Any) -> int:
    return sum(
        1
        for row in range(3, (ws.max_row or 0) + 1)
        if _row_has_content(ws, row, len(V2_HEADERS))
    )


def main(argv: list[str] | None = None) -> int:
    configure_utf8_stdio()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", nargs="?", type=Path, help="要校验的 .xlsx 文件路径")
    parser.add_argument("--path", dest="opt_path", type=Path, help="要校验的 .xlsx 文件路径")
    parser.add_argument(
        "--sheet-name", default="测试用例", help="主工作表名称（默认为“测试用例”）"
    )
    args = parser.parse_args(argv)

    target = args.opt_path or args.path
    if target is None:
        parser.error("必须提供 .xlsx 文件路径")

    errors = validate_excel_file(target, sheet_name=args.sheet_name)
    if errors:
        print(f"Excel 校验失败: 发现 {len(errors)} 处错误:", file=sys.stderr)
        for index, error in enumerate(errors, start=1):
            print(f"{index}. {error}", file=sys.stderr)
        return 1

    print(f"Excel 校验通过: {target}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
