#!/usr/bin/env python3
"""Validate structure, headers, mandatory fields, and uniqueness for Excel (.xlsx) test cases."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from traceability_matrix import (
    configure_utf8_stderr,
    entries_for_case_id,
    final_matrix_errors,
    load_traceability_matrix,
    TRACEABILITY_PATTERN,
    traceability_text,
)
from validation_utils import is_valid_date_or_placeholder

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
    "用例 ID",
    "用例标题",
    "用例类型",
    "设计方法",
    "优先级",
    "前置条件",
    "执行步骤",
    "预期结果",
    "备注",
]

ALLOWED_PRIORITIES = ("P0", "P1", "P2", "P3", "P4")
ALLOWED_CASE_TYPES = (
    "功能测试",
    "回归测试",
    "性能测试",
    "安全测试",
    "兼容性测试",
    "易用性测试",
    "可靠性测试",
)
ALLOWED_DESIGN_METHODS = (
    "等价类划分",
    "边界值分析",
    "判定表法",
    "状态迁移法",
    "错误推测法",
    "正交试验",
    "场景分析法",
)

REQUIREMENTS_FILE = Path(__file__).resolve().parents[1] / "requirements.txt"
INFO_LABELS = ("需求名称", "上线日期", "用例数量", "覆盖范围", "补充说明")
REQUIRED_INFO_LABELS = ("需求名称", "上线日期", "用例数量")
COUNT_PATTERN = re.compile(r"[0-9]+")


def _cell_text(value: object) -> str:
    return str(value).strip() if value is not None else ""


def _row_has_content(ws: object, row_idx: int) -> bool:
    values = [
        ws.cell(row=row_idx, column=column).value
        for column in range(1, len(EXPECTED_HEADERS) + 1)
    ]
    return any(value is not None and str(value).strip() for value in values)


def count_case_rows(ws: object) -> int:
    return sum(
        1
        for row_idx in range(3, (ws.max_row or 0) + 1)
        if _row_has_content(ws, row_idx)
    )


def last_case_row(ws: object) -> int:
    indexes = [
        row_idx
        for row_idx in range(3, (ws.max_row or 0) + 1)
        if _row_has_content(ws, row_idx)
    ]
    return indexes[-1] if indexes else 2


def validate_info_sheet(workbook: object, expected_count: int) -> list[str]:
    if "生成说明" not in workbook.sheetnames:
        return ["工作簿中缺少名为 '生成说明' 的工作表。"]

    info = workbook["生成说明"]
    errors: list[str] = []
    label_rows: dict[str, list[int]] = {label: [] for label in INFO_LABELS}
    for row_idx in range(1, (info.max_row or 0) + 1):
        label = _cell_text(info.cell(row=row_idx, column=1).value)
        if label in label_rows:
            label_rows[label].append(row_idx)

    values: dict[str, str] = {}
    for label in INFO_LABELS:
        rows = label_rows[label]
        if not rows:
            errors.append(f"生成说明缺少字段“{label}”。")
        elif len(rows) > 1:
            errors.append(f"生成说明字段“{label}”重复出现 {len(rows)} 次。")
        else:
            values[label] = _cell_text(info.cell(row=rows[0], column=2).value)

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

    for row_idx in range(1, (info.max_row or 0) + 1):
        for column in (1, 2):
            if info.cell(row=row_idx, column=column).data_type == "f":
                errors.append(f"生成说明第 {row_idx} 行包含公式，必须使用纯文本。")
    for row_idx in range(1, (info.max_row or 0) + 1):
        for column in range(3, (info.max_column or 0) + 1):
            value = info.cell(row=row_idx, column=column).value
            if value is not None and str(value).strip():
                errors.append(
                    f"生成说明第 {row_idx} 行第 {column} 列存在额外内容，只允许 A–B 两列。"
                )
                break
    return errors


def validate_excel_file(
    file_path: Path,
    traceability_matrix: dict[str, dict[str, object]] | None = None,
    sheet_name: str = "测试用例",
) -> list[str]:
    errors: list[str] = []

    if not file_path.is_file():
        return [f"文件不存在: {file_path}"]
    if traceability_matrix is None:
        return ["未提供需求追溯矩阵，无法核对追溯是否来自材料。"]
    matrix_errors = final_matrix_errors(traceability_matrix)
    errors.extend(matrix_errors)
    matrix_finalized = not matrix_errors

    try:
        from openpyxl import load_workbook
    except ImportError:
        return [
            "缺少 openpyxl 依赖，无法读取和校验 .xlsx 文件，请执行: "
            f'pip install -r "{REQUIREMENTS_FILE}"'
        ]

    try:
        wb = load_workbook(filename=str(file_path), data_only=False)
    except Exception as exc:
        return [f"无法读取 Excel 文件: {exc}"]

    if sheet_name not in wb.sheetnames:
        return [f"工作簿中缺少名为 '{sheet_name}' 的工作表，当前包含的工作表: {wb.sheetnames}"]

    ws = wb[sheet_name]
    max_row = ws.max_row or 0
    actual_case_rows = count_case_rows(ws)
    max_column = ws.max_column or 0

    a1_text = _cell_text(ws["A1"].value)
    if not a1_text:
        errors.append("工作表的 A1 标题不能为空。")
    if ws.cell(row=1, column=1).data_type == "f":
        errors.append("第 1 行标题包含公式，必须使用纯文本。")
    merge_refs = {str(merged) for merged in ws.merged_cells.ranges}
    if merge_refs != {"A1:K1"}:
        errors.append("工作表必须恰好包含 A1:K1 合并标题区域，不得存在其他合并区域。")

    if max_column > len(EXPECTED_HEADERS):
        for row_idx in range(1, max_row + 1):
            for column in range(len(EXPECTED_HEADERS) + 1, max_column + 1):
                value = ws.cell(row=row_idx, column=column).value
                if value is not None and str(value).strip():
                    errors.append(
                        f"第 {row_idx} 行第 {column} 列存在额外内容，只允许 A–K 共 11 列。"
                    )
                    break

    errors.extend(validate_info_sheet(wb, actual_case_rows))

    # 检查第 2 行为列头
    row2_values = [ws.cell(row=2, column=col).value for col in range(1, len(EXPECTED_HEADERS) + 1)]
    row2_cleaned = [str(v).strip() if v is not None else "" for v in row2_values]

    if row2_cleaned != EXPECTED_HEADERS:
        errors.append(
            f"第 2 行表头不符合规范。\n期望: {EXPECTED_HEADERS}\n实际: {row2_cleaned}"
        )
    for col_idx in range(1, len(EXPECTED_HEADERS) + 1):
        if ws.cell(row=2, column=col_idx).data_type == "f":
            errors.append(
                f"第 2 行字段“{EXPECTED_HEADERS[col_idx - 1]}”包含公式，必须使用纯文本。"
            )

    if max_row < 3:
        errors.append("工作表中未发现有效用例数据行（数据应从第 3 行开始）。")
        return errors

    # 列索引映射
    header_to_col = {header: idx + 1 for idx, header in enumerate(EXPECTED_HEADERS)}

    seen_case_ids: set[str] = set()
    seen_titles: set[str] = set()
    actual_case_ids: set[str] = set()

    if ws.freeze_panes != "A3":
        errors.append("工作表必须冻结窗格至 A3。")
    expected_filter_ref = f"A2:K{last_case_row(ws)}"
    if ws.auto_filter.ref != expected_filter_ref:
        errors.append(
            f"工作表筛选范围必须为 {expected_filter_ref}，实际为 {ws.auto_filter.ref or '未设置'}。"
        )

    for row_idx in range(3, max_row + 1):
        # 检查该行是否全空
        row_vals = [ws.cell(row=row_idx, column=c).value for c in range(1, len(EXPECTED_HEADERS) + 1)]
        if not _row_has_content(ws, row_idx):
            continue  # 跳过全空行

        for col_idx in range(1, len(EXPECTED_HEADERS) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.data_type == "f":
                errors.append(
                    f"第 {row_idx} 行字段“{EXPECTED_HEADERS[col_idx - 1]}”包含公式，必须使用纯文本。"
                )

        # 检查必填项
        for field in MANDATORY_FIELDS:
            col_idx = header_to_col[field]
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is None or not str(val).strip():
                errors.append(f"第 {row_idx} 行缺少必填字段: '{field}'。")

        if any(
            "[待确认]" in str(value or "")
            for value in row_vals
        ):
            errors.append(f"第 {row_idx} 行不得包含旧版“[待确认]”标记。")
        notes_val = ws.cell(row=row_idx, column=header_to_col["备注"]).value
        notes = str(notes_val).strip() if notes_val is not None else ""
        note_lines = notes.splitlines()
        traceability_lines: list[str] = []
        for note_line in note_lines:
            traceability = TRACEABILITY_PATTERN.match(note_line)
            if not traceability:
                break
            if not all(part.strip() for part in traceability.groups()):
                break
            traceability_lines.append(note_line)
        if not traceability_lines:
            errors.append(
                f"第 {row_idx} 行备注开头必须为“[需求追溯] 来源：<材料>｜定位：<位置>｜规则：<规则>”，不得仅使用内部编号。"
            )
        elif any(line.startswith("[需求追溯]") for line in note_lines[len(traceability_lines):]):
            errors.append(f"第 {row_idx} 行的多条需求追溯必须连续输出。")

        case_id = str(ws.cell(row=row_idx, column=header_to_col["用例 ID"]).value or "").strip()
        if case_id:
            actual_case_ids.add(case_id)
            matched_entries = entries_for_case_id(traceability_matrix, case_id)
            expected_traceability = [
                traceability_text(entry)
                for entry in matched_entries
            ]
            if matrix_finalized:
                if not matched_entries:
                    errors.append(
                        f"第 {row_idx} 行用例 ID “{case_id}”未在需求追溯矩阵中关联。"
                    )
                elif traceability_lines != expected_traceability:
                    errors.append(f"第 {row_idx} 行备注中的需求追溯与矩阵记录不一致。")

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
            p_str = str(priority_val).strip()
            if p_str not in ALLOWED_PRIORITIES:
                errors.append(
                    f"第 {row_idx} 行优先级无效: '{priority_val}'，必须为 {'、'.join(ALLOWED_PRIORITIES)} 之一。"
                )

        case_type = str(ws.cell(row=row_idx, column=header_to_col["用例类型"]).value or "").strip()
        if case_type and case_type not in ALLOWED_CASE_TYPES:
            errors.append(
                f"第 {row_idx} 行用例类型无效: '{case_type}'，必须为 {'、'.join(ALLOWED_CASE_TYPES)} 之一。"
            )

        design_method = str(ws.cell(row=row_idx, column=header_to_col["设计方法"]).value or "").strip()
        if design_method and design_method not in ALLOWED_DESIGN_METHODS:
            errors.append(
                f"第 {row_idx} 行设计方法无效: '{design_method}'，必须为 {'、'.join(ALLOWED_DESIGN_METHODS)} 之一。"
            )

    if actual_case_rows == 0:
        errors.append("工作表中未包含任何有效用例数据行。")
    if matrix_finalized:
        for entry in traceability_matrix.values():
            if entry["status"] != "已覆盖":
                continue
            if not entry["case_ids"]:
                errors.append(
                    f"需求追溯矩阵中的已覆盖条目“{entry['id']}”缺少 Excel 用例 ID 关联。"
                )
            for case_id in entry["case_ids"]:
                if case_id not in actual_case_ids:
                    errors.append(
                        f"需求追溯矩阵中的 Excel 用例 ID “{case_id}”未出现在工作表中。"
                    )

    return errors


def main() -> None:
    configure_utf8_stderr()
    parser = argparse.ArgumentParser(description="校验 Excel (.xlsx) 测试用例文件的结构、表头与必填字段。")
    parser.add_argument("path", nargs="?", help="要校验的 .xlsx 文件路径")
    parser.add_argument("--path", dest="opt_path", help="要校验的 .xlsx 文件路径（兼容选项）")
    parser.add_argument("--sheet-name", default="测试用例", help="测试用例工作表名称（默认为'测试用例'）")
    parser.add_argument("--traceability-matrix", required=True, help="需求追溯矩阵 JSON 文件路径")
    args = parser.parse_args()

    target_path_str = args.opt_path or args.path
    if not target_path_str:
        parser.error("必须提供 .xlsx 文件路径")

    target_path = Path(target_path_str).expanduser().resolve()
    traceability_matrix, matrix_errors = load_traceability_matrix(args.traceability_matrix)
    errors = matrix_errors or validate_excel_file(
        target_path,
        traceability_matrix,
        sheet_name=args.sheet_name,
    )

    if errors:
        print(f"校验失败: 发现 {len(errors)} 处错误:\n", file=sys.stderr)
        for idx, err in enumerate(errors, start=1):
            print(f"{idx}. {err}", file=sys.stderr)
        sys.exit(1)

    print(f"校验通过: {target_path}")
    sys.exit(0)


if __name__ == "__main__":
    main()
