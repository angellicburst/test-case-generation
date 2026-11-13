#!/usr/bin/env python3
"""Validate list rendering for Markdown test-case fields."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

FIELD_PATTERN = re.compile(r"^(\s*)- (前置条件|测试步骤|预期结果)\s*$")
CHILD_PATTERN = re.compile(r"^(\s{2,})(?:(-)|(\d+)\.)\s+.+$")
LEGACY_CASE_PATTERN = re.compile(r"^####\s+.+-P[0-4]\s*$")
CHECKBOX_CASE_PATTERN = re.compile(r"^\s*- \[ \] (正向|反向)-(.+)-P[0-4]\s*$")
CHECKED_CASE_PATTERN = re.compile(r"^\s*- \[[xX]\] (?:正向|反向)-.+-P[0-4]\s*$")
CASE_WITH_ID_PATTERN = re.compile(
    r"^\s*- \[[ xX]\]\s*(?:正向|反向)-TC(?:-[^-]+)?-\d+-.+-P[0-4]\s*$"
)
CASE_ITEM_PATTERN = re.compile(r"^\s*- \[[ xX]\]\s+([^-]+?)-.*-P[0-4]\s*$")
NODE_PATTERN = re.compile(r"^(\s*)- \[[ xX]\]\s+")
DESCRIPTION_PATTERN = re.compile(r"^\s*>\s*用例数量[:：]\s*用例\s*(\d+)\s*条\s*$")
EXPECTED_RESULT_PATTERN = re.compile(r"^\s*-\s*预期结果\s*$")


def validate(text: str) -> list[str]:
    lines = [line.expandtabs(4) for line in text.splitlines()]
    errors: list[str] = []
    case_titles: list[tuple[int, str]] = []
    declared_count: int | None = None
    expected_result_count = 0
    node_stack: list[int] = []

    for index, line in enumerate(lines):
        description = DESCRIPTION_PATTERN.match(line)
        if description:
            declared_count = int(description.group(1))
        if EXPECTED_RESULT_PATTERN.match(line):
            expected_result_count += 1
        case_item = CASE_ITEM_PATTERN.match(line)
        if case_item and case_item.group(1) not in ("正向", "反向"):
            errors.append(
                f"第 {index + 1} 行的用例类型非法，必须为 正向 或 反向，实际为 '{case_item.group(1)}'。"
            )
        if LEGACY_CASE_PATTERN.match(line):
            errors.append(
                f"第 {index + 1} 行的用例必须使用 `- [ ] 用例类型-用例标题-优先级` 列表项，不得使用 `####` 标题。"
            )
            continue
        if CASE_WITH_ID_PATTERN.match(line):
            errors.append(
                f"第 {index + 1} 行的用例任务项不得包含用例编号，Markdown 格式已移除用例编号。"
            )
        if CHECKED_CASE_PATTERN.match(line):
            errors.append(f"第 {index + 1} 行的用例任务项不得使用已勾选状态 `- [x]`。")

        node = NODE_PATTERN.match(line)
        if not node:
            continue
        indent = len(node.group(1))
        while node_stack and node_stack[-1] >= indent:
            node_stack.pop()
        checkbox_case = CHECKBOX_CASE_PATTERN.match(line)
        if checkbox_case and not CASE_WITH_ID_PATTERN.match(line):
            case_titles.append((index + 1, checkbox_case.group(2).strip()))
            if len(node_stack) < 2:
                errors.append(
                    f"第 {index + 1} 行的用例必须位于功能模块和测试类型的列表项下（至少两级缩进）。"
                )
        node_stack.append(indent)

    seen_titles: set[str] = set()
    for line_number, case_title in case_titles:
        if case_title in seen_titles:
            errors.append(f"第 {line_number} 行的用例标题“{case_title}”重复。")
        seen_titles.add(case_title)

    if declared_count is None:
        errors.append("文档顶部缺少用例数量声明（`> 用例数量: 用例 N 条`）。")
    elif declared_count != expected_result_count:
        errors.append(
            f"文档声明用例数量为 {declared_count}，实际“预期结果”数量为 {expected_result_count}。"
        )

    for index, line in enumerate(lines):
        match = FIELD_PATTERN.match(line)
        if not match:
            continue

        field_indent = len(match.group(1))
        field = match.group(2)
        items: list[tuple[int, str, int | None]] = []
        cursor = index + 1
        while cursor < len(lines):
            current = lines[cursor]
            current_indent = len(current) - len(current.lstrip(" "))
            if current.strip() and current_indent <= field_indent and (current.lstrip().startswith("- ") or current.lstrip().startswith("####")):
                break
            child = CHILD_PATTERN.match(current)
            if child and len(child.group(1)) > field_indent:
                items.append(
                    (
                        cursor + 1,
                        "unordered" if child.group(2) else "ordered",
                        int(child.group(3)) if child.group(3) else None,
                    )
                )
            cursor += 1

        if not items:
            errors.append(f"第 {index + 1} 行的“{field}”缺少内容列表。")
            continue

        if len(items) == 1:
            if items[0][1] != "unordered":
                errors.append(
                    f"第 {items[0][0]} 行的“{field}”仅有一项，必须使用无序列表。"
                )
            continue

        if any(kind != "ordered" for _, kind, _ in items):
            errors.append(
                f"第 {index + 1} 行的“{field}”有 {len(items)} 项，必须全部使用有序列表。"
            )
            continue

        expected = list(range(1, len(items) + 1))
        actual = [number for _, _, number in items]
        if actual != expected:
            errors.append(
                f"第 {index + 1} 行的“{field}”有序列表必须从 1 开始连续编号，实际为 {actual}。"
            )

    return errors


def fix(text: str) -> str:
    lines = text.splitlines(keepends=True)
    index = 0
    while index < len(lines):
        if not FIELD_PATTERN.match(lines[index].rstrip("\r\n")):
            index += 1
            continue

        field_indent = len(FIELD_PATTERN.match(lines[index].rstrip("\r\n")).group(1))
        item_indexes: list[int] = []
        cursor = index + 1
        while cursor < len(lines):
            current = lines[cursor].rstrip("\r\n")
            current_indent = len(current) - len(current.lstrip(" "))
            if current.strip() and current_indent <= field_indent and (current.lstrip().startswith("- ") or current.lstrip().startswith("####")):
                break
            child = CHILD_PATTERN.match(current)
            if child and len(child.group(1)) > field_indent:
                item_indexes.append(cursor)
            cursor += 1

        if len(item_indexes) > 1:
            for number, item_index in enumerate(item_indexes, start=1):
                newline = "\r\n" if lines[item_index].endswith("\r\n") else "\n"
                item = lines[item_index].rstrip("\r\n")
                indent = re.match(r"^(\s*)", item).group(1)
                content = re.sub(r"^\s*(?:(?:-)|(?:\d+\.))\s+", "", item)
                lines[item_index] = f"{indent}{number}. {content}{newline}"
        index = cursor

    return "".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", type=Path, help="Markdown test-case file to validate")
    parser.add_argument(
        "--fix",
        action="store_true",
        help="Convert multi-item field lists to ordered lists",
    )
    args = parser.parse_args()

    if not args.path.is_file():
        print(f"文件不存在: {args.path}", file=sys.stderr)
        return 1

    try:
        text = args.path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError as exc:
        print(f"文件编码不是 UTF-8，无法校验: {args.path}（{exc}）", file=sys.stderr)
        return 1

    if args.fix:
        fixed = fix(text)
        if fixed != text:
            args.path.write_text(fixed, encoding="utf-8")
            text = fixed

    errors = validate(text)
    if errors:
        print("Markdown list validation failed:", file=sys.stderr)
        print("\n".join(f"- {error}" for error in errors), file=sys.stderr)
        return 1

    print(f"Markdown list validation passed: {args.path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
