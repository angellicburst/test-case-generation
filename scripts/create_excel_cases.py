#!/usr/bin/env python3
"""Generate a test case workbook in .xlsx format using the standard 11-column template."""

from __future__ import annotations

import argparse
import json
import os
import sys
import tempfile
from pathlib import Path
from typing import Any

from case_baseline import load_baseline, normalize_baseline, validate_baseline
from traceability_matrix import load_traceability_matrix, traceability_text
from validation_utils import (
    FileLockedError,
    configure_utf8_stdio,
    estimate_row_height as estimate_text_row_height,
    format_excel_multiline,
    is_file_locked_error,
    is_valid_date_or_placeholder,
)

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

REQUIREMENTS_FILE = Path(__file__).resolve().parents[1] / "requirements.txt"

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


def as_excel_text(value: Any) -> str:
    return str(value) if value is not None else ""


def normalize_multiline(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, list):
        return "\n".join(str(item) for item in value if item is not None)
    return str(value)


def build_notes(notes_value: Any, traceability_lines: list[str]) -> str:
    notes = normalize_multiline(notes_value).strip()
    if notes.startswith("[需求追溯]"):
        raise ValueError("备注不得手工填写需求追溯，请使用 traceability_refs。")
    return "\n".join([*traceability_lines, *([notes] if notes else [])])


def traceability_entries_for_case(
    case: dict[str, Any],
    traceability_matrix: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    raw_refs = case.get("traceability_refs")
    if not isinstance(raw_refs, list) or not raw_refs:
        raise ValueError("每条用例必须提供非空的 traceability_refs 数组。")

    refs = [str(reference).strip() for reference in raw_refs]
    if any(not reference for reference in refs) or len(set(refs)) != len(refs):
        raise ValueError("traceability_refs 必须是无重复的非空追溯 ID。")

    case_id = str(case.get("id", case.get("case_id", ""))).strip()
    entries_by_reference: dict[str, dict[str, Any]] = {}
    for reference in refs:
        entry = traceability_matrix.get(reference)
        if entry is None:
            raise ValueError(f"用例 '{case_id or '未命名用例'}' 包含无效的 traceability_refs: '{reference}'。")
        if entry["status"] != "已覆盖":
            raise ValueError(f"用例 '{case_id}' 只能关联已覆盖的追溯项。")
        if case_id not in entry["case_ids"]:
            raise ValueError(
                f"用例 '{case_id}' 不在追溯项 '{reference}' 的 case_ids 中。"
            )
        entries_by_reference[reference] = entry
    return [
        entry
        for trace_id, entry in traceability_matrix.items()
        if trace_id in entries_by_reference
    ]


def estimate_row_height(values: list[Any]) -> int:
    longest = max((len(str(value or "")) for value in values), default=0)
    line_breaks = sum(str(value or "").count("\n") for value in values)
    estimated_lines = max(1, line_breaks + 1, longest // 32 + 1)
    return min(max(24, estimated_lines * 18), 180)


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
            normalize_multiline(meta.get("notes") or meta.get("generated_note", "")),
        ),
    ]

    thin = Side(style="thin", color="B7C3D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    label_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")

    for row_index, (label, value) in enumerate(rows, start=2):
        label_cell = sheet.cell(row=row_index, column=1, value=label)
        value_cell = sheet.cell(row=row_index, column=2, value=as_excel_text(value))
        value_cell.data_type = "s"
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


def build_case_sheet(
    workbook: Any,
    title: str,
    cases: list[dict[str, Any]],
    traceability_matrix: dict[str, dict[str, Any]],
    sheet_name: str = "测试用例",
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

    for idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=2, column=idx, value=header)
        cell.font = Font(name="Microsoft YaHei", bold=True, size=11)
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 24

    for row_index, case in enumerate(cases, start=3):
        traceability_entries = traceability_entries_for_case(case, traceability_matrix)
        traceability_lines = [traceability_text(entry) for entry in traceability_entries]
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
            build_notes(case.get("notes", case.get("remark", "")), traceability_lines),
        ]
        for col_index, value in enumerate(values, start=1):
            cell = ws.cell(row=row_index, column=col_index, value=as_excel_text(value))
            cell.font = Font(name="Microsoft YaHei", size=10)
            cell.data_type = "s"
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


def validate_meta(meta: dict[str, Any] | None) -> dict[str, Any]:
    if meta is None:
        meta = {}
    if not isinstance(meta, dict):
        raise ValueError("'meta' 必须是对象")
    if "case_count" in meta:
        raise ValueError("meta.case_count 不得填写，用例数量按 cases 数组长度自动计算。")

    normalized = dict(meta)
    requirement_name = str(normalized.get("requirement_name") or "").strip()
    if not requirement_name:
        raise ValueError("meta.requirement_name 必须为非空字符串。")
    normalized["requirement_name"] = requirement_name

    release_date = str(normalized.get("release_date") or "").strip()
    if not release_date:
        raise ValueError("meta.release_date 必须为非空字符串。")
    if not is_valid_date_or_placeholder(release_date):
        raise ValueError("meta.release_date 必须为有效 YYYY-MM-DD 或 未指定日期。")
    normalized["release_date"] = release_date
    return normalized


def load_cases_payload(input_json: str | Path) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    input_path = Path(input_json)
    try:
        payload_text = input_path.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as exc:
        raise ValueError(f"无法读取输入 JSON: {input_path}（{exc}）") from None
    try:
        payload = json.loads(payload_text)
    except json.JSONDecodeError as exc:
        raise ValueError(f"输入 JSON 格式错误: {input_path}（{exc}）") from None
    if not isinstance(payload, dict):
        raise ValueError("输入 JSON 必须为包含 'cases' 数组的对象")

    cases = payload.get("cases", [])
    if not isinstance(cases, list):
        raise ValueError("'cases' 必须是数组")
    if not cases:
        raise ValueError("'cases' 必须是非空数组")

    meta = payload.get("meta", {})
    if meta is not None and not isinstance(meta, dict):
        raise ValueError("'meta' 必须是对象")

    return validate_meta(meta), cases


def build_workbook(
    title: str,
    cases: list[dict[str, Any]],
    traceability_matrix: dict[str, dict[str, Any]],
    sheet_name: str = "测试用例",
    meta: dict[str, Any] | None = None,
) -> Any:
    meta = validate_meta(meta)
    if not cases:
        raise ValueError("'cases' 必须是非空数组")
    title_text = str(title).strip()
    if not title_text:
        raise ValueError("标题不能为空。")
    meta["case_count"] = len(cases)

    try:
        from openpyxl import Workbook
    except ImportError as exc:
        raise RuntimeError(
            f'生成 xlsx 需要 openpyxl 库，请先安装: pip install -r "{REQUIREMENTS_FILE}"'
        ) from exc

    wb = Workbook()
    build_case_sheet(wb, title_text, cases, traceability_matrix, sheet_name)
    build_info_sheet(wb, meta)
    return wb


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
        if is_file_locked_error(exc):
            raise FileLockedError(
                f"文件被占用，无法写入：{path}。"
                "请先关闭正在打开该文件的程序（如 Excel、WPS 或编辑器）后重试。"
            ) from exc
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
    for row_index, case in enumerate(cases, start=3):
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
    ws.auto_filter.ref = f"A2:K{max(len(cases) + 2, 2)}"


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
    resolved_sheet = sheet_name or meta["sheet_name"]
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


def main() -> None:
    configure_utf8_stdio()
    parser = argparse.ArgumentParser(description="生成标准的 .xlsx 格式测试用例工作簿。")
    parser.add_argument("--output", required=True, help="输出 .xlsx 文件绝对路径")
    parser.add_argument("--baseline", default=None, help="合并 IR JSON 文件路径（推荐）")
    parser.add_argument(
        "--title",
        default=None,
        help="[旧版输入] 第 1 行显示的大标题；省略时按 meta 的上线日期与需求名称生成",
    )
    parser.add_argument("--sheet-name", default=None, help="主工作表名称")
    parser.add_argument(
        "--input-json",
        default=None,
        help="合并 IR JSON 路径（兼容别名）；与 --traceability-matrix 同用时表示旧版输入",
    )
    parser.add_argument(
        "--traceability-matrix",
        default=None,
        help="[已废弃] 旧版需求追溯矩阵 JSON 路径",
    )
    parser.add_argument("--force", action="store_true", help="输出文件已存在时确认覆盖")
    args = parser.parse_args()

    try:
        output_path = Path(args.output).expanduser().resolve()
        if output_path.exists() and output_path.is_dir():
            raise SystemExit("输出路径必须是文件路径，不能是目录")
        if output_path.exists() and not args.force:
            raise SystemExit(
                f"输出文件已存在: {output_path}。确认覆盖后请追加 --force。"
            )
        output_path.parent.mkdir(parents=True, exist_ok=True)

        use_legacy = bool(args.input_json and args.traceability_matrix)
        if use_legacy:
            print(
                "警告：--input-json + --traceability-matrix 为旧版输入方式，请改用 --baseline。",
                file=sys.stderr,
            )
            meta, cases = load_cases_payload(args.input_json)
            traceability_matrix, matrix_errors = load_traceability_matrix(
                args.traceability_matrix
            )
            if matrix_errors:
                raise SystemExit("\n".join(matrix_errors))
            title = (
                meta.get("title")
                or args.title
                or f"测试用例 - {meta['release_date']} - [{meta['requirement_name']}]"
            )
            title_text = str(title).strip()
            if not title_text:
                raise ValueError("标题不能为空。")
            sheet_name = meta.get("sheet_name") or args.sheet_name or "测试用例"
            workbook = build_workbook(
                title_text,
                cases,
                traceability_matrix,
                sheet_name=sheet_name,
                meta=meta,
            )
        else:
            source = args.baseline or args.input_json
            if not source:
                raise SystemExit("请通过 --baseline 指定合并 IR 文件。")
            workbook = build_workbook_from_ir(load_baseline(source), sheet_name=args.sheet_name)

        save_workbook_atomic(workbook, output_path)
    except ValueError as exc:
        raise SystemExit(str(exc)) from None
    print(f"成功生成测试用例工作簿: {output_path}")


if __name__ == "__main__":
    main()

