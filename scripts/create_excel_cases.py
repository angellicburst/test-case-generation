#!/usr/bin/env python3
"""Generate a test case workbook in .xlsx format from the merged IR."""

from __future__ import annotations

import argparse
import os
import sys
import tempfile
from pathlib import Path
from typing import Any

from case_baseline import (
    canonical_tree_order,
    load_baseline_raw,
    normalize_baseline,
    prepare_render_ir,
    validate_baseline,
)
from validation_utils import (
    configure_utf8_stdio,
    estimate_row_height as estimate_text_row_height,
    format_excel_multiline,
    is_file_locked_error,
    locked_write_error,
    readonly_write_error,
)

REQUIREMENTS_FILE = Path(__file__).resolve().parents[1] / "requirements.txt"

V2_HEADERS = (
    "追溯 ID",
    "用例 ID",
    "用例标题",
    "用例类型",
    "设计方法",
    "优先级",
    "前置操作",
    "测试数据",
    "测试步骤",
    "预期结果",
    "后置操作",
)
V2_COLUMN_WIDTHS = [22, 18, 36, 12, 18, 10, 32, 26, 44, 44, 32]
TRACE_HEADERS = ("追溯 ID", "来源", "定位", "规则", "关联用例 ID")
TRACE_COLUMN_WIDTHS = [20, 28, 28, 64, 28]
PENDING_HEADERS = ("追溯 ID", "来源", "定位", "未明确项", "当前测试假设", "确认问题")
PENDING_COLUMN_WIDTHS = [20, 26, 26, 32, 32, 36]
RESERVED_SHEET_NAMES = ("生成说明", "需求追溯", "待确认项")
INFO_LABELS = ("需求名称", "上线日期", "用例数量", "覆盖范围", "补充说明")


def as_excel_text(value: Any) -> str:
    return str(value) if value is not None else ""


def normalize_multiline(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, list):
        return "\n".join(str(item) for item in value if item is not None)
    return str(value)


def build_info_sheet(workbook: Any, meta: dict[str, Any]) -> None:
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    sheet = workbook.create_sheet("生成说明")
    sheet["A1"] = "测试用例生成说明"
    sheet.merge_cells("A1:B1")
    sheet["A1"].font = Font(name="Microsoft YaHei", bold=True, size=14)
    sheet["A1"].alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 26

    rows = [
        ("需求名称", meta.get("requirement_name", "")),
        ("上线日期", meta.get("release_date", "")),
        ("用例数量", str(meta.get("case_count", ""))),
        ("覆盖范围", normalize_multiline(meta.get("coverage", ""))),
        (
            "补充说明",
            normalize_multiline(meta.get("notes", "")),
        ),
    ]

    thin = Side(style="thin", color="B7C3D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    label_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")

    for row_index, (label, value) in enumerate(rows, start=2):
        label_cell = sheet.cell(row=row_index, column=1, value=label)
        value_text = as_excel_text(value)
        value_cell = sheet.cell(row=row_index, column=2, value=value_text)
        value_cell.data_type = "s"
        label_cell.font = Font(name="Microsoft YaHei", bold=True, size=10)
        value_cell.font = Font(name="Microsoft YaHei", size=10)
        label_cell.fill = label_fill
        label_cell.border = border
        value_cell.border = border
        label_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        value_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        sheet.row_dimensions[row_index].height = estimate_text_row_height(
            [label, value_text], [22, 80]
        )

    sheet.column_dimensions["A"].width = 22
    sheet.column_dimensions["B"].width = 80


def save_workbook_atomic(workbook: Any, path: Path) -> None:
    fd, temp_name = tempfile.mkstemp(
        prefix=f".{path.stem}.",
        suffix=path.suffix or ".xlsx",
        dir=str(path.parent),
    )
    os.close(fd)
    try:
        workbook.save(temp_name)
        os.replace(temp_name, path)
    except OSError as exc:
        try:
            os.unlink(temp_name)
        except OSError:
            pass
        readonly_error = readonly_write_error(exc, path)
        if readonly_error is not None:
            raise readonly_error from exc
        if is_file_locked_error(exc) or getattr(exc, "winerror", None) == 5:
            raise locked_write_error(exc, path) from exc
        raise
    except BaseException:
        try:
            os.unlink(temp_name)
        except OSError:
            pass
        raise


def build_case_sheet_v2(
    workbook: Any,
    title: str,
    cases: list[dict[str, Any]],
    traceability: list[dict[str, Any]],
    sheet_name: str,
) -> None:
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    ws = workbook.active
    ws.title = sheet_name

    ws["A1"] = as_excel_text(title)
    ws["A1"].data_type = "s"
    ws.merge_cells("A1:K1")
    ws["A1"].font = Font(name="Microsoft YaHei", bold=True, size=16)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    thin = Side(style="thin", color="B7C3D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")

    for index, header in enumerate(V2_HEADERS, start=1):
        cell = ws.cell(row=2, column=index, value=header)
        cell.font = Font(name="Microsoft YaHei", bold=True, size=11)
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 24

    trace_order = {entry["id"]: index for index, entry in enumerate(traceability)}
    default_order = len(trace_order)
    ordered_cases = canonical_tree_order(cases)
    for row_index, case in enumerate(ordered_cases, start=3):
        refs = sorted(
            case["traceability_refs"], key=lambda ref: trace_order.get(ref, default_order)
        )
        values = [
            ", ".join(refs),
            case["id"],
            case["title"],
            case["type"],
            case["design_method"],
            case["priority"],
            format_excel_multiline(case["pre_operations"]),
            format_excel_multiline(case["test_data"]),
            format_excel_multiline(case["steps"]),
            format_excel_multiline(case["expected_results"]),
            format_excel_multiline(case["post_operations"]),
        ]
        for column, value in enumerate(values, start=1):
            cell = ws.cell(row=row_index, column=column, value=as_excel_text(value))
            cell.font = Font(name="Microsoft YaHei", size=10)
            cell.data_type = "s"
            cell.border = border
            horizontal = "center" if column in (4, 5, 6) else "left"
            cell.alignment = Alignment(horizontal=horizontal, vertical="top", wrap_text=True)
        ws.row_dimensions[row_index].height = estimate_text_row_height(
            values, V2_COLUMN_WIDTHS
        )

    for index, width in enumerate(V2_COLUMN_WIDTHS, start=1):
        ws.column_dimensions[chr(64 + index)].width = width

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:K{max(len(ordered_cases) + 2, 2)}"


def build_traceability_sheet(workbook: Any, entries: list[dict[str, Any]]) -> None:
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    sheet = workbook.create_sheet("需求追溯")
    sheet["A1"] = "需求追溯表"
    sheet.merge_cells("A1:E1")
    sheet["A1"].font = Font(name="Microsoft YaHei", bold=True, size=14)
    sheet["A1"].alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 26

    thin = Side(style="thin", color="B7C3D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    for index, header in enumerate(TRACE_HEADERS, start=1):
        cell = sheet.cell(row=2, column=index, value=header)
        cell.font = Font(name="Microsoft YaHei", bold=True, size=11)
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.row_dimensions[2].height = 24

    for row_index, entry in enumerate(entries, start=3):
        values = [
            entry["id"],
            entry["source"],
            entry["location"],
            entry["rule"],
            ", ".join(entry.get("case_ids", [])),
        ]
        for column, value in enumerate(values, start=1):
            cell = sheet.cell(row=row_index, column=column, value=as_excel_text(value))
            cell.font = Font(name="Microsoft YaHei", size=10)
            cell.data_type = "s"
            cell.border = border
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        sheet.row_dimensions[row_index].height = estimate_text_row_height(
            values, TRACE_COLUMN_WIDTHS
        )

    for index, width in enumerate(TRACE_COLUMN_WIDTHS, start=1):
        sheet.column_dimensions[chr(64 + index)].width = width
    sheet.freeze_panes = "A3"
    sheet.auto_filter.ref = f"A2:E{max(len(entries) + 2, 2)}"


def build_pending_sheet(workbook: Any, entries: list[dict[str, Any]]) -> None:
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    sheet = workbook.create_sheet("待确认项")
    sheet["A1"] = "待确认项清单"
    sheet.merge_cells("A1:F1")
    sheet["A1"].font = Font(name="Microsoft YaHei", bold=True, size=14)
    sheet["A1"].alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 26

    thin = Side(style="thin", color="B7C3D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    for index, header in enumerate(PENDING_HEADERS, start=1):
        cell = sheet.cell(row=2, column=index, value=header)
        cell.font = Font(name="Microsoft YaHei", bold=True, size=11)
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.row_dimensions[2].height = 24

    for row_index, entry in enumerate(entries, start=3):
        values = [
            entry["id"],
            entry["source"],
            entry["location"],
            entry.get("issue", ""),
            entry.get("assumption", ""),
            entry.get("question", ""),
        ]
        for column, value in enumerate(values, start=1):
            cell = sheet.cell(row=row_index, column=column, value=as_excel_text(value))
            cell.font = Font(name="Microsoft YaHei", size=10)
            cell.data_type = "s"
            cell.border = border
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        sheet.row_dimensions[row_index].height = estimate_text_row_height(
            values, PENDING_COLUMN_WIDTHS
        )

    for index, width in enumerate(PENDING_COLUMN_WIDTHS, start=1):
        sheet.column_dimensions[chr(64 + index)].width = width
    sheet.freeze_panes = "A3"
    sheet.auto_filter.ref = f"A2:F{max(len(entries) + 2, 2)}"


def build_workbook_from_ir(ir: dict[str, Any], sheet_name: str | None = None) -> Any:
    errors = validate_baseline(ir, require_design_method=True)
    if errors:
        raise ValueError("\n".join(errors))

    payload = normalize_baseline(ir)
    meta = payload["meta"]
    cases = payload["cases"]
    entries = payload["traceability"]
    resolved_sheet = (
        str(sheet_name).strip() if sheet_name else ""
    ) or meta["sheet_name"]
    if resolved_sheet in RESERVED_SHEET_NAMES:
        raise ValueError(f"工作表名不得使用保留名: {'、'.join(RESERVED_SHEET_NAMES)}。")

    try:
        from openpyxl import Workbook
    except ImportError as exc:
        raise RuntimeError(
            f'生成 xlsx 需要 openpyxl 库，请先安装: pip install -r "{REQUIREMENTS_FILE}"'
        ) from exc

    workbook = Workbook()
    build_case_sheet_v2(workbook, meta["title"], cases, entries, resolved_sheet)
    build_info_sheet(workbook, {**meta, "case_count": len(cases)})
    build_traceability_sheet(
        workbook, [entry for entry in entries if entry["status"] != "待确认"]
    )
    pending = [entry for entry in entries if entry["status"] == "待确认"]
    if pending:
        build_pending_sheet(workbook, pending)
    return workbook


def main(argv: list[str] | None = None) -> int:
    configure_utf8_stdio()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True, help="输出 .xlsx 文件路径")
    parser.add_argument("--baseline", required=True, help="合并 IR JSON 文件路径")
    parser.add_argument("--sheet-name", default=None, help="主工作表名称")
    parser.add_argument("--force", action="store_true", help="输出文件已存在时确认覆盖")
    args = parser.parse_args(argv)

    output_path = Path(args.output).expanduser().resolve()
    if output_path.exists() and output_path.is_dir():
        print("输出路径必须是文件路径，不能是目录", file=sys.stderr)
        return 1
    if output_path.exists() and not args.force:
        print(
            f"输出文件已存在: {output_path}。确认覆盖后请追加 --force。", file=sys.stderr
        )
        return 1

    try:
        raw = load_baseline_raw(args.baseline)
        ir, errors, warnings = prepare_render_ir(
            raw, require_design_method=True
        )
        if errors:
            print("\n".join(errors), file=sys.stderr)
            return 1
        for warning in warnings:
            print(f"提示：{warning}", file=sys.stderr)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        workbook = build_workbook_from_ir(ir, sheet_name=args.sheet_name)
        save_workbook_atomic(workbook, output_path)
    except (ValueError, OSError, RuntimeError) as exc:
        print(str(exc), file=sys.stderr)
        return 1

    print(f"成功生成测试用例工作簿: {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
