#!/usr/bin/env python3
"""Generate a test case workbook in .xlsx format using the standard 11-column template."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

HEADERS = [
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

DEFAULT_COLUMN_WIDTHS = {
    "A": 18,
    "B": 18,
    "C": 36,
    "D": 12,
    "E": 18,
    "F": 10,
    "G": 32,
    "H": 26,
    "I": 44,
    "J": 44,
    "K": 26,
}


def normalize_multiline(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, list):
        return "\n".join(str(item) for item in value if item is not None)
    return str(value)


def estimate_row_height(values: list[Any]) -> int:
    longest = max((len(str(value or "")) for value in values), default=0)
    line_breaks = sum(str(value or "").count("\n") for value in values)
    estimated_lines = max(1, line_breaks + 1, longest // 32 + 1)
    return min(max(24, estimated_lines * 18), 180)


def build_info_sheet(workbook: Any, meta: dict[str, Any]) -> None:
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    sheet = workbook.create_sheet("生成说明")
    sheet["A1"] = "测试用例生成与需求存疑说明"
    sheet.merge_cells("A1:B1")
    sheet["A1"].font = Font(name="Microsoft YaHei", bold=True, size=14)
    sheet["A1"].alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 26

    rows = [
        ("需求名称", meta.get("requirement_name") or meta.get("title", "")),
        ("上线日期", meta.get("release_date", "")),
        ("用例数量", str(meta.get("case_count", ""))),
        ("覆盖范围", normalize_multiline(meta.get("coverage", ""))),
        ("需求存疑与待确认清单", normalize_multiline(meta.get("open_questions", "无待确认项"))),
        ("补充说明", meta.get("notes") or meta.get("generated_note", "")),
    ]

    thin = Side(style="thin", color="B7C3D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    label_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")

    for row_index, (label, value) in enumerate(rows, start=2):
        label_cell = sheet.cell(row=row_index, column=1, value=label)
        value_cell = sheet.cell(row=row_index, column=2, value=value)
        label_cell.font = Font(name="Microsoft YaHei", bold=True, size=10)
        value_cell.font = Font(name="Microsoft YaHei", size=10)
        label_cell.fill = label_fill
        label_cell.border = border
        value_cell.border = border
        label_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        value_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        sheet.row_dimensions[row_index].height = estimate_row_height([value])

    sheet.column_dimensions["A"].width = 22
    sheet.column_dimensions["B"].width = 80


def build_case_sheet(workbook: Any, title: str, cases: list[dict[str, Any]], sheet_name: str = "测试用例") -> None:
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    ws = workbook.active
    ws.title = sheet_name

    ws["A1"] = title
    ws.merge_cells("A1:K1")
    ws["A1"].font = Font(name="Microsoft YaHei", bold=True, size=16)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    thin = Side(style="thin", color="B7C3D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")

    for idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=2, column=idx, value=header)
        cell.font = Font(name="Microsoft YaHei", bold=True, size=11)
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 24

    for row_index, case in enumerate(cases, start=3):
        values = [
            case.get("requirement_id", case.get("trace_id", "")),
            case.get("id", case.get("case_id", "")),
            case.get("title", case.get("case_title", "")),
            case.get("type", case.get("case_type", "")),
            case.get("design_method", ""),
            case.get("priority", ""),
            normalize_multiline(case.get("preconditions", case.get("precondition", ""))),
            normalize_multiline(case.get("test_data", "")),
            normalize_multiline(case.get("steps", case.get("step_description", ""))),
            normalize_multiline(case.get("expected_results", case.get("expected_result", ""))),
            normalize_multiline(case.get("notes", case.get("remark", ""))),
        ]
        for col_index, value in enumerate(values, start=1):
            cell = ws.cell(row=row_index, column=col_index, value=value)
            cell.font = Font(name="Microsoft YaHei", size=10)
            cell.border = border
            if col_index in (4, 5, 6):  # 类型、方法、优先级居中
                cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        ws.row_dimensions[row_index].height = estimate_row_height(values)

    for col, width in DEFAULT_COLUMN_WIDTHS.items():
        ws.column_dimensions[col].width = width

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:K{max(len(cases) + 2, 2)}"


def load_cases_payload(input_json: str | Path) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    payload = json.loads(Path(input_json).read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("输入 JSON 必须为包含 'cases' 数组的对象")

    cases = payload.get("cases", [])
    if not isinstance(cases, list):
        raise ValueError("'cases' 必须是数组")

    meta = payload.get("meta", {})
    if meta is not None and not isinstance(meta, dict):
        raise ValueError("'meta' 必须是对象")

    return meta or {}, cases


def build_workbook(title: str, cases: list[dict[str, Any]], sheet_name: str = "测试用例", meta: dict[str, Any] | None = None) -> Any:
    try:
        from openpyxl import Workbook
    except ImportError as exc:
        raise RuntimeError("生成 xlsx 需要 openpyxl 库，请先安装: pip install openpyxl") from exc

    wb = Workbook()
    build_case_sheet(wb, title, cases, sheet_name)
    meta = meta or {}
    if "case_count" not in meta:
        meta["case_count"] = len(cases)
    build_info_sheet(wb, meta)
    return wb


def main() -> None:
    parser = argparse.ArgumentParser(description="生成标准的 .xlsx 格式测试用例工作簿。")
    parser.add_argument("--output", required=True, help="输出 .xlsx 文件绝对路径")
    parser.add_argument("--title", default="测试用例表", help="第 1 行显示的大标题")
    parser.add_argument("--sheet-name", default="测试用例", help="主工作表名称")
    parser.add_argument("--input-json", required=True, help="包含 cases 列表与可选 meta 的 JSON 文件路径")
    args = parser.parse_args()

    meta, cases = load_cases_payload(args.input_json)
    title = meta.get("title") or args.title
    sheet_name = meta.get("sheet_name") or args.sheet_name

    output_path = Path(args.output).expanduser().resolve()
    if output_path.exists() and output_path.is_dir():
        raise SystemExit("输出路径必须是文件路径，不能是目录")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = build_workbook(title, cases, sheet_name=sheet_name, meta=meta)
    workbook.save(output_path)
    print(f"成功生成测试用例工作簿: {output_path}")


if __name__ == "__main__":
    main()

