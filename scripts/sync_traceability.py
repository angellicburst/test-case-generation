#!/usr/bin/env python3
"""Synchronize Markdown and Excel traceability text from a traceability matrix."""

from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any

from traceability_matrix import (
    configure_utf8_stderr,
    entries_for_case_id,
    entries_for_case_title,
    load_traceability_matrix,
    TRACEABILITY_PATTERN,
    traceability_text,
)

MARKDOWN_CASE_PATTERN = re.compile(
    r"^(\s*)- \[[ xX]\]\s+(正向|反向)-(.+)-P[0-4]\s*$"
)
MARKDOWN_QUOTE_PATTERN = re.compile(r"^(\s*)>\s*(\S.*)$")


def traceability_lines(entries: list[dict[str, Any]]) -> list[str]:
    return [traceability_text(entry) for entry in entries]


def sync_markdown(
    path: Path,
    matrix: dict[str, dict[str, Any]],
) -> list[str]:
    try:
        with open(path, encoding="utf-8-sig", newline="") as handle:
            lines = handle.read().splitlines(keepends=True)
    except UnicodeDecodeError as exc:
        return [f"Markdown 文件编码不是 UTF-8: {path}（{exc}）"]
    except OSError as exc:
        return [f"无法读取 Markdown 文件: {path}（{exc}）"]

    errors: list[str] = []
    index = 0
    while index < len(lines):
        match = MARKDOWN_CASE_PATTERN.match(lines[index].rstrip("\r\n"))
        if not match:
            index += 1
            continue

        case_indent = len(match.group(1))
        case_title = match.group(3).strip()
        entries = entries_for_case_title(matrix, case_title)
        if not entries:
            errors.append(f"Markdown 用例标题“{case_title}”未在需求追溯矩阵中关联。")
            index += 1
            continue

        cursor = index + 1
        while cursor < len(lines) and not lines[cursor].strip():
            cursor += 1
        if cursor < len(lines):
            first_quote = MARKDOWN_QUOTE_PATTERN.match(lines[cursor].rstrip("\r\n"))
            if (
                first_quote
                and len(first_quote.group(1)) == case_indent + 2
                and not TRACEABILITY_PATTERN.match(first_quote.group(2).strip())
            ):
                errors.append(
                    f"Markdown 用例标题“{case_title}”下首个引用块不是需求追溯，未执行同步。"
                )
                index += 1
                continue
        trace_start = cursor
        while cursor < len(lines):
            trace_match = MARKDOWN_QUOTE_PATTERN.match(lines[cursor].rstrip("\r\n"))
            if (
                not trace_match
                or len(trace_match.group(1)) != case_indent + 2
                or not TRACEABILITY_PATTERN.match(trace_match.group(2).strip())
            ):
                break
            cursor += 1

        newline = "\r\n" if lines[index].endswith("\r\n") else "\n"
        prefix = " " * (case_indent + 2)
        replacement = [
            f"{prefix}> {traceability}{newline}"
            for traceability in traceability_lines(entries)
        ]
        lines[trace_start:cursor] = replacement
        index = trace_start + len(replacement)

    if errors:
        return errors
    with open(path, "w", encoding="utf-8", newline="") as handle:
        handle.write("".join(lines))
    return []


def sync_excel(
    path: Path,
    matrix: dict[str, dict[str, Any]],
    sheet_name: str,
) -> list[str]:
    try:
        from openpyxl import load_workbook
    except ImportError:
        return ["缺少 openpyxl 依赖，无法同步 Excel 需求追溯。"]

    try:
        workbook = load_workbook(path)
    except Exception as exc:
        return [f"无法读取 Excel 文件: {path}（{exc}）"]
    if sheet_name not in workbook.sheetnames:
        return [f"工作簿中缺少名为 '{sheet_name}' 的工作表。"]

    sheet = workbook[sheet_name]
    errors: list[str] = []
    for row_index in range(3, (sheet.max_row or 0) + 1):
        values = [sheet.cell(row=row_index, column=column).value for column in range(1, 12)]
        if not any(value is not None and str(value).strip() for value in values):
            continue
        case_id = str(sheet.cell(row=row_index, column=2).value or "").strip()
        entries = entries_for_case_id(matrix, case_id)
        if not entries:
            errors.append(f"第 {row_index} 行 Excel 用例 ID“{case_id}”未在需求追溯矩阵中关联。")
            continue

        notes = str(sheet.cell(row=row_index, column=11).value or "")
        note_lines = notes.splitlines()
        cursor = 0
        while cursor < len(note_lines) and note_lines[cursor].startswith("[需求追溯]"):
            cursor += 1
        notes_after_traceability = "\n".join(note_lines[cursor:]).strip()
        value = "\n".join(
            [*traceability_lines(entries), *([notes_after_traceability] if notes_after_traceability else [])]
        )
        cell = sheet.cell(row=row_index, column=11, value=value)
        cell.data_type = "s"

    if errors:
        return errors
    workbook.save(path)
    return []


def _copy_to_temp(path: Path) -> Path:
    fd, temp_name = tempfile.mkstemp(
        prefix=f".{path.stem}.",
        suffix=path.suffix or ".tmp",
        dir=str(path.parent),
    )
    os.close(fd)
    try:
        shutil.copy2(path, temp_name)
    except BaseException:
        try:
            os.unlink(temp_name)
        except OSError:
            pass
        raise
    return Path(temp_name)


def sync_atomically(
    markdown_path: Path | None,
    excel_path: Path | None,
    matrix: dict[str, dict[str, Any]],
    sheet_name: str = "测试用例",
) -> list[str]:
    targets: list[tuple[Path, str]] = []
    if markdown_path is not None:
        targets.append((markdown_path, "Markdown"))
    if excel_path is not None:
        targets.append((excel_path, "Excel"))
    for path, label in targets:
        if not path.is_file():
            return [f"{label} 文件不存在: {path}"]
    labels = dict(targets)

    staged: dict[Path, Path] = {}
    backups: dict[Path, Path] = {}
    backups_to_keep: set[Path] = set()
    try:
        try:
            for path, _label in targets:
                staged[path] = _copy_to_temp(path)
        except OSError as exc:
            return [f"无法创建待写入的临时副本: {exc}"]

        errors: list[str] = []
        if markdown_path is not None:
            errors.extend(sync_markdown(staged[markdown_path], matrix))
        if not errors and excel_path is not None:
            errors.extend(sync_excel(staged[excel_path], matrix, sheet_name))
        if errors:
            return errors

        try:
            for path, _label in targets:
                backups[path] = _copy_to_temp(path)
        except OSError as exc:
            return [f"无法创建原文件备份，未写入任何正式文件: {exc}"]

        committed: list[Path] = []
        commit_order = [path for path in (excel_path, markdown_path) if path is not None]
        for path in commit_order:
            try:
                os.replace(staged[path], path)
                staged.pop(path)
            except OSError as exc:
                rollback_errors: list[str] = []
                for committed_path in reversed(committed):
                    backup_path = backups.pop(committed_path)
                    try:
                        os.replace(backup_path, committed_path)
                    except OSError as rollback_exc:
                        backups_to_keep.add(backup_path)
                        rollback_errors.append(
                            f"回滚 {labels[committed_path]} 失败: {rollback_exc}；"
                            f"原文件备份保留在 {backup_path}，请手工恢复。"
                        )
                message = f"写入 {labels[path]} 失败: {path}（{exc}）。"
                if rollback_errors:
                    message += "部分已写入文件回滚失败，请按下列备份路径手工恢复。"
                else:
                    message += "已恢复此前写入的文件。"
                return [message, *rollback_errors]
            committed.append(path)
        return []
    finally:
        for temp_path in [*staged.values(), *backups.values()]:
            if temp_path in backups_to_keep:
                continue
            try:
                temp_path.unlink()
            except OSError:
                pass


def main() -> int:
    configure_utf8_stderr()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--traceability-matrix", type=Path, required=True)
    parser.add_argument("--markdown", type=Path)
    parser.add_argument("--excel", type=Path)
    parser.add_argument("--sheet-name", default="测试用例")
    args = parser.parse_args()

    if args.markdown is None and args.excel is None:
        parser.error("至少提供 --markdown 或 --excel 之一。")

    matrix, errors = load_traceability_matrix(args.traceability_matrix)
    if not errors:
        errors.extend(sync_atomically(args.markdown, args.excel, matrix, args.sheet_name))

    if errors:
        print("追溯同步失败:", file=sys.stderr)
        for error in errors:
            print(f"- {error}", file=sys.stderr)
        return 1

    print("需求追溯同步完成。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
