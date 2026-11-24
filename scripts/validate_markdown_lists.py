#!/usr/bin/env python3
"""Validate a self-contained Markdown v2 test-case document."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Any

from case_baseline import (
    describe_ir_differences,
    load_baseline,
    parse_markdown_text,
    validate_baseline,
)
from validation_utils import FileLockedError, configure_utf8_stdio, write_text_atomic

FIELD_HEADER_PATTERN = re.compile(
    r"^( *)- (前置操作|测试数据|测试步骤|预期结果|后置操作)$"
)
ITEM_PATTERN = re.compile(r"^( *)(?:(-)|(\d+)\.) +(.*)$")
HEADING_PATTERN = re.compile(r"^##\s")


def _list_format_errors(lines: list[str]) -> list[str]:
    """Check list markers: one item unordered, several items ordered from 1."""
    errors: list[str] = []
    index = 0
    while index < len(lines):
        header = FIELD_HEADER_PATTERN.match(lines[index])
        if not header:
            index += 1
            continue
        header_index = index
        field_indent = len(header.group(1))
        field = header.group(2)
        items: list[tuple[int, int | None]] = []
        cursor = index + 1
        while cursor < len(lines):
            line = lines[cursor]
            if not line.strip():
                cursor += 1
                continue
            indent = len(line) - len(line.lstrip(" "))
            if indent <= field_indent:
                break
            item = ITEM_PATTERN.match(line)
            if indent == field_indent + 2 and item:
                items.append((cursor, None if item.group(2) else int(item.group(3))))
            cursor += 1
        index = cursor

        if len(items) <= 1:
            if items and items[0][1] is not None:
                errors.append(
                    f"第 {items[0][0] + 1} 行：“{field}”仅有一项时必须使用无序列表 `- `。"
                )
            continue
        if any(number is None for _, number in items):
            errors.append(
                f"第 {header_index + 1} 行：“{field}”有 {len(items)} 项，必须全部使用有序列表 `1.`。"
            )
            continue
        numbers = [number for _, number in items]
        if numbers != list(range(1, len(items) + 1)):
            errors.append(
                f"第 {header_index + 1} 行：“{field}”的有序列表必须从 1 开始连续编号，"
                f"实际为 {numbers}。"
            )
    return errors


def validate(text: str, baseline: dict[str, Any] | None = None) -> list[str]:
    """Validate one Markdown document without any external traceability matrix."""
    ir, parse_errors, _warnings = parse_markdown_text(text)
    errors = list(parse_errors)
    if errors:
        return list(dict.fromkeys(errors))

    errors.extend(validate_baseline(ir))
    errors.extend(_list_format_errors(text.splitlines()))
    if baseline is not None:
        errors.extend(
            describe_ir_differences(
                baseline, ir, "基线 IR", "Markdown", include_design_method=False
            )
        )
    return list(dict.fromkeys(errors))


def fix(text: str) -> str:
    """Renumber multi-item field lists as ordered lists; leave the trace section alone."""
    lines = text.splitlines(keepends=True)
    index = 0
    while index < len(lines):
        raw = lines[index].rstrip("\r\n")
        if HEADING_PATTERN.match(raw):
            break
        header = FIELD_HEADER_PATTERN.match(raw)
        if not header:
            index += 1
            continue

        field_indent = len(header.group(1))
        item_indexes: list[int] = []
        cursor = index + 1
        while cursor < len(lines):
            current = lines[cursor].rstrip("\r\n")
            if not current.strip():
                cursor += 1
                continue
            indent = len(current) - len(current.lstrip(" "))
            if indent <= field_indent:
                break
            item = ITEM_PATTERN.match(current)
            if indent == field_indent + 2 and item:
                item_indexes.append(cursor)
            cursor += 1

        if len(item_indexes) > 1:
            for number, item_index in enumerate(item_indexes, start=1):
                newline = "\r\n" if lines[item_index].endswith("\r\n") else "\n"
                item = lines[item_index].rstrip("\r\n")
                indent = re.match(r"^ *", item).group(0)
                content = ITEM_PATTERN.match(item).group(4)
                lines[item_index] = f"{indent}{number}. {content}{newline}"
        index = cursor

    return "".join(lines)


def main(argv: list[str] | None = None) -> int:
    configure_utf8_stdio()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", type=Path, help="Markdown 测试用例文件路径")
    parser.add_argument(
        "--fix",
        action="store_true",
        help="把多项字段列表改写为有序列表（不改动需求追溯表）",
    )
    parser.add_argument("--baseline", type=Path, help="可选的合并 IR，用于比对差异")
    args = parser.parse_args(argv)

    if not args.path.is_file():
        print(f"文件不存在: {args.path}", file=sys.stderr)
        return 1

    try:
        with args.path.open(encoding="utf-8-sig", newline="") as handle:
            text = handle.read()
    except (OSError, UnicodeDecodeError) as exc:
        print(f"无法读取 Markdown 文件: {args.path}（{exc}）", file=sys.stderr)
        return 1

    baseline = None
    if args.baseline is not None:
        try:
            baseline = load_baseline(args.baseline)
        except ValueError as exc:
            print(str(exc), file=sys.stderr)
            return 1

    if args.fix:
        fixed = fix(text)
        if fixed != text:
            try:
                write_text_atomic(args.path, fixed)
            except (OSError, FileLockedError) as exc:
                print(str(exc), file=sys.stderr)
                return 1
            text = fixed

    errors = validate(text, baseline)
    if errors:
        print(f"Markdown 校验失败: 发现 {len(errors)} 处错误:", file=sys.stderr)
        for index, error in enumerate(errors, start=1):
            print(f"{index}. {error}", file=sys.stderr)
        return 1

    cases = text.count("- [ ] 正向-") + text.count("- [ ] 反向-")
    print(f"Markdown 校验通过: {args.path}（{cases} 条用例）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
