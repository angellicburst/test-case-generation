#!/usr/bin/env python3
"""Validate list rendering for Markdown test-case fields."""

from __future__ import annotations

import argparse
import os
import re
import sys
import tempfile
from collections import Counter
from datetime import date
from pathlib import Path

from traceability_matrix import (
    configure_utf8_stderr,
    entries_for_case_title,
    final_matrix_errors,
    load_traceability_matrix,
    TRACEABILITY_PATTERN,
    traceability_text,
)

FIELD_PATTERN = re.compile(r"^(\s*)- (前置条件|测试步骤|预期结果)\s*$")
TRACEABILITY_QUOTE_PATTERN = re.compile(r"^(\s*)>\s*(\S.*)$")
CHILD_PATTERN = re.compile(r"^(\s{2,})(?:(-)|(\d+)\.)\s+.+$")
LEGACY_CASE_PATTERN = re.compile(r"^####\s+.+-P[0-4]\s*$")
CHECKBOX_CASE_PATTERN = re.compile(r"^\s*- \[ \] (正向|反向)-(.+)-P[0-4]\s*$")
CHECKED_CASE_PATTERN = re.compile(r"^\s*- \[[xX]\] (?:正向|反向)-.+-P[0-4]\s*$")
CASE_WITH_ID_PATTERN = re.compile(
    r"^\s*- \[[ xX]\]\s*(?:正向|反向)-TC(?:-[^-]+)?-\d+-.+-P[0-4]\s*$"
)
CASE_ITEM_PATTERN = re.compile(r"^\s*- \[[ xX]\]\s+([^-]+?)-.*-P[0-4]\s*$")
NODE_PATTERN = re.compile(r"^(\s*)- \[[ xX]\]\s+")
LIST_ITEM_PATTERN = re.compile(r"^(\s*)(?:-|\d+\.)\s+\S.*$")
DESCRIPTION_PATTERN = re.compile(r"^\s*>\s*用例数量[:：]\s*用例\s*(\d+)\s*条\s*$")
TITLE_PATTERN = re.compile(r"^# 测试用例 - (?:\d{4}-\d{2}-\d{2}|未指定日期) - \[.+\]\s*$")
TITLE_DATE_PATTERN = re.compile(r"^# 测试用例 - (\S+) - \[")
ALLOWED_TEST_TYPES = frozenset(
    {
        "功能测试",
        "回归测试",
        "性能测试",
        "安全测试",
        "兼容性测试",
        "易用性测试",
        "可靠性测试",
    }
)


def validate(
    text: str,
    traceability_matrix: dict[str, dict[str, object]] | None = None,
) -> list[str]:
    lines = [line.expandtabs(4) for line in text.splitlines()]
    errors: list[str] = []
    cases: list[tuple[int, str, int]] = []
    declared_count: int | None = None
    matrix_finalized = False
    node_stack: list[int] = []
    description_indexes: list[int] = []

    if not lines or not TITLE_PATTERN.match(lines[0]):
        errors.append("文档首行必须为 `# 测试用例 - YYYY-MM-DD - [需求名称]`（无具体日期时用 `未指定日期`）。")
    else:
        title_date = TITLE_DATE_PATTERN.match(lines[0]).group(1)
        if title_date != "未指定日期":
            try:
                date.fromisoformat(title_date)
            except ValueError:
                errors.append(
                    f"文档标题中的日期“{title_date}”不是有效日历日期，必须为 YYYY-MM-DD。"
                )
        first_content_after_title = next(
            (index for index, line in enumerate(lines[1:], start=1) if line.strip()),
            None,
        )
        if first_content_after_title is None or not DESCRIPTION_PATTERN.match(
            lines[first_content_after_title]
        ):
            errors.append("文档标题后首个非空行必须为用例数量声明。")
        else:
            declared_count = int(
                DESCRIPTION_PATTERN.match(lines[first_content_after_title]).group(1)
            )

    for index, line in enumerate(lines):
        description = DESCRIPTION_PATTERN.match(line)
        if description:
            description_indexes.append(index)
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
        if indent not in (0, 2, 4, 6):
            errors.append(f"第 {index + 1} 行的任务清单缩进必须为 2 个空格的层级。")
        elif indent and (not node_stack or node_stack[-1] != indent - 2):
            errors.append(f"第 {index + 1} 行的任务清单必须紧邻上一级节点。")
        if indent == 2:
            node_label = line[node.end():].strip()
            if node_label not in ALLOWED_TEST_TYPES:
                errors.append(
                    f"第 {index + 1} 行的测试类型非法，必须为"
                    "功能测试、回归测试、性能测试、安全测试、兼容性测试、易用性测试、可靠性测试之一，"
                    f"实际为 '{node_label}'。"
                )
        if checkbox_case and not CASE_WITH_ID_PATTERN.match(line):
            cases.append((index, checkbox_case.group(2).strip(), indent))
            if indent not in (4, 6) or len(node_stack) < 2:
                errors.append(
                    f"第 {index + 1} 行的用例必须位于功能模块和测试类型的列表项下，缩进为 4 或 6 个空格。"
                )
        elif indent == 6:
            errors.append(f"第 {index + 1} 行的 6 空格层级只能用于用例任务项。")
        node_stack.append(indent)

    seen_titles: set[str] = set()
    for index, case_title, _ in cases:
        line_number = index + 1
        if case_title in seen_titles:
            errors.append(f"第 {line_number} 行的用例标题“{case_title}”重复。")
        seen_titles.add(case_title)

    if not cases:
        errors.append("文档中未包含任何有效用例任务项。")
    if traceability_matrix is None:
        errors.append("未提供需求追溯矩阵，无法核对追溯是否来自材料。")
    else:
        matrix_errors = final_matrix_errors(traceability_matrix)
        errors.extend(matrix_errors)
        matrix_finalized = not matrix_errors
    if len(description_indexes) > 1:
        errors.append("文档只能包含一个用例数量声明。")
    if declared_count is None:
        errors.append("文档顶部缺少用例数量声明（`> 用例数量: 用例 N 条`）。")
    elif declared_count != len(cases):
        errors.append(
            f"文档声明用例数量为 {declared_count}，实际用例任务项数量为 {len(cases)}。"
        )

    actual_case_titles: set[str] = set()
    for case_position, (case_index, case_title, case_indent) in enumerate(cases):
        actual_case_titles.add(case_title)
        case_end = cases[case_position + 1][0] if case_position + 1 < len(cases) else len(lines)
        field_names: list[str] = []
        traceability_indexes: list[int] = []
        actual_traceability: list[str] = []
        first_content_index = next(
            (
                content_index
                for content_index, content_line in enumerate(lines[case_index + 1:case_end], start=case_index + 1)
                if content_line.strip()
            ),
            None,
        )
        for content_index, line in enumerate(lines[case_index + 1:case_end], start=case_index + 1):
            node = NODE_PATTERN.match(line)
            if node and len(node.group(1)) <= case_indent:
                break
            field = FIELD_PATTERN.match(line)
            if field and len(field.group(1)) == case_indent + 2:
                field_names.append(field.group(2))
            direct_item = LIST_ITEM_PATTERN.match(line)
            if (
                direct_item
                and len(direct_item.group(1)) == case_indent + 2
                and not field
            ):
                direct_content = re.sub(r"^\s*(?:-|\d+\.)\s+", "", line).strip()
                errors.append(
                    f"第 {content_index + 1} 行的用例只允许直接包含“前置条件”“测试步骤”“预期结果”字段，"
                    f"发现非法字段“{direct_content}”。"
                )
            elif (
                line.strip()
                and len(line) - len(line.lstrip(" ")) == case_indent + 2
                and not field
                and not TRACEABILITY_QUOTE_PATTERN.match(line)
            ):
                errors.append(
                    f"第 {content_index + 1} 行的用例只允许直接包含需求追溯引用以及"
                    f"“前置条件”“测试步骤”“预期结果”字段，发现非法直接内容“{line.strip()}”。"
                )
            traceability = TRACEABILITY_QUOTE_PATTERN.match(line)
            if traceability and len(traceability.group(1)) == case_indent + 2:
                traceability_indexes.append(content_index)
                traceability_content = TRACEABILITY_PATTERN.match(traceability.group(2).strip())
                if not traceability_content or not all(part.strip() for part in traceability_content.groups()):
                    errors.append(
                        f"第 {content_index + 1} 行的需求追溯必须填写“来源、定位、规则”，不得仅使用内部编号。"
                    )
                else:
                    actual_traceability.append(traceability.group(2).strip())
        field_counts = Counter(field_names)
        if not traceability_indexes:
            errors.append(f"第 {case_index + 1} 行的用例标题下缺少需求追溯引用。")
        else:
            expected_indexes = list(
                range(traceability_indexes[0], traceability_indexes[0] + len(traceability_indexes))
            )
            if traceability_indexes[0] != first_content_index:
                errors.append(f"第 {case_index + 1} 行的用例必须在标题下首个非空行填写需求追溯。")
            if traceability_indexes != expected_indexes:
                errors.append(f"第 {case_index + 1} 行的多条需求追溯必须连续输出。")
            if traceability_matrix is not None and matrix_finalized:
                matched_entries = entries_for_case_title(traceability_matrix, case_title)
                expected_traceability = [
                    traceability_text(entry)
                    for entry in matched_entries
                ]
                if not matched_entries:
                    errors.append(
                        f"第 {case_index + 1} 行的用例标题“{case_title}”未在需求追溯矩阵中关联。"
                    )
                elif actual_traceability != expected_traceability:
                    errors.append(
                        f"第 {case_index + 1} 行的需求追溯与矩阵记录不一致。"
                    )
        if "[待确认]" in "\n".join(lines[case_index:case_end]):
            errors.append(f"第 {case_index + 1} 行的用例不得包含旧版“[待确认]”标记。")
        for field in ("前置条件", "测试步骤", "预期结果"):
            if field_counts[field] == 0:
                errors.append(f"第 {case_index + 1} 行的用例缺少必填字段“{field}”。")
            elif field_counts[field] > 1:
                errors.append(f"第 {case_index + 1} 行的用例包含重复的“{field}”字段。")

    if traceability_matrix is not None and matrix_finalized:
        for entry in traceability_matrix.values():
            if entry["status"] != "已覆盖":
                continue
            for case_title in entry["case_titles"]:
                if case_title not in actual_case_titles:
                    errors.append(
                        f"需求追溯矩阵中的 Markdown 用例标题“{case_title}”未出现在文档中。"
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
            if child and len(child.group(1)) == field_indent + 2:
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
            if child and len(child.group(1)) == field_indent + 2:
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


def write_text_atomic(path: Path, text: str) -> None:
    fd, temp_name = tempfile.mkstemp(
        prefix=f".{path.name}.",
        suffix=".tmp",
        dir=str(path.parent),
    )
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            handle.write(text)
        os.replace(temp_name, path)
    except BaseException:
        try:
            os.unlink(temp_name)
        except OSError:
            pass
        raise


def main() -> int:
    configure_utf8_stderr()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", type=Path, help="Markdown test-case file to validate")
    parser.add_argument(
        "--traceability-matrix",
        type=Path,
        required=True,
        help="需求追溯矩阵 JSON 文件路径",
    )
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
            write_text_atomic(args.path, fixed)
            text = fixed

    traceability_matrix, matrix_errors = load_traceability_matrix(args.traceability_matrix)
    errors = matrix_errors or validate(text, traceability_matrix)
    if errors:
        print("Markdown list validation failed:", file=sys.stderr)
        print("\n".join(f"- {error}" for error in errors), file=sys.stderr)
        return 1

    print(f"Markdown list validation passed: {args.path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
