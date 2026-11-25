#!/usr/bin/env python3
"""Load, validate, and round-trip the merged test case IR."""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path
from typing import Any

from traceability_matrix import (
    describe_traceability_mismatch,
    parse_trace_id_list,
    sync_traceability_from_cases,
    validate_traceability_entries,
)
from validation_utils import (
    ALLOWED_CASE_TYPES,
    ALLOWED_DESIGN_METHODS,
    ALLOWED_POLARITIES,
    ALLOWED_PRIORITIES,
    FileLockedError,
    configure_utf8_stdio,
    is_valid_date_or_placeholder,
    normalize_list_items,
    split_md_table_row,
    write_text_atomic,
)

SCHEMA_VERSION = "2.0"
RESERVED_SHEET_NAMES = ("生成说明", "需求追溯", "待确认项")
PLACEHOLDER_IDS = {"", "TC-NEW", "NEW"}
ID_PATTERN = re.compile(r"^(.+?)-(\d+)$")
ID_INVALID_PATTERN = re.compile(r"[\s,\[\]，]")
TITLE_PATTERN = re.compile(r"^测试用例\s*-\s*(.+?)\s*-\s*\[(.+)\]\s*$")
COUNT_PATTERN = re.compile(r"^>\s*用例数量[:：]\s*用例\s*(\d+)\s*条\s*$")
CASE_LINE_PATTERN = re.compile(r"^(\s*)- \[ \] (正向|反向)-(.+)-(P[0-4])\s*$")
CASE_ID_IN_TITLE_PATTERN = re.compile(r"TC-[A-Z0-9]+-\d+")
REF_LINE_PATTERN = re.compile(r"^\s*> (?:\[([^\[\]]+)\])?\s*追溯[:：]\s*(.+?)\s*$")
FIELD_PATTERN = re.compile(r"^- (前置操作|测试数据|测试步骤|预期结果|后置操作)$")
ITEM_PATTERN = re.compile(r"^(?:[-*]|\d+\.)\s+(.*)$")
NODE_PATTERN = re.compile(r"^- \[ \] (.+)$")
TOKEN_PATTERN = re.compile(r"^:?-{3,}:?$")

FIELD_ORDER = ("前置操作", "测试数据", "测试步骤", "预期结果", "后置操作")
FIELD_KEY_MAP = {
    "前置操作": "pre_operations",
    "测试数据": "test_data",
    "测试步骤": "steps",
    "预期结果": "expected_results",
    "后置操作": "post_operations",
}
REQUIRED_FIELDS = ("前置操作", "测试步骤", "预期结果")
UNSUPPORTED_FIELDS = ("备注", "前置条件", "后置清理")

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
COVERED_HEADERS = ("追溯 ID", "来源", "定位", "规则", "关联用例")
COVERED_HEADERS_ID = ("追溯 ID", "来源", "定位", "规则", "关联用例 ID")
PENDING_HEADERS = ("追溯 ID", "来源", "定位", "未明确项", "当前测试假设", "确认问题")
INFO_LABELS = ("需求名称", "上线日期", "用例数量", "覆盖范围", "补充说明")


def _text(value: Any) -> str:
    return str(value).strip() if value is not None else ""


def _multiline(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (list, tuple)):
        return "\n".join(str(item) for item in value if item is not None).strip()
    return str(value).strip()


def _has_newline(value: str) -> bool:
    return "\r" in value or "\n" in value


def _contains_marker(value: Any) -> bool:
    if isinstance(value, (list, tuple)):
        return any(_contains_marker(item) for item in value)
    return "[待确认]" in str(value)


def normalize_meta(raw: Any) -> dict[str, Any]:
    meta = dict(raw) if isinstance(raw, dict) else {}
    requirement_name = _text(meta.get("requirement_name"))
    release_date = _text(meta.get("release_date"))
    title = _text(meta.get("title"))
    if not title and requirement_name and release_date:
        title = f"测试用例 - {release_date} - [{requirement_name}]"
    return {
        "title": title,
        "requirement_name": requirement_name,
        "release_date": release_date,
        "coverage": _multiline(meta.get("coverage")),
        "notes": _multiline(meta.get("notes")),
        "sheet_name": _text(meta.get("sheet_name")) or "测试用例",
    }


def normalize_case(raw: Any) -> dict[str, Any]:
    case = raw if isinstance(raw, dict) else {}
    refs = case.get("traceability_refs")
    if isinstance(refs, (list, tuple)):
        normalized_refs: Any = [value for value in (_text(item) for item in refs) if value]
    elif refs is None:
        normalized_refs = []
    else:
        normalized_refs = refs
    return {
        "id": _text(case.get("id")),
        "module": _text(case.get("module")),
        "type": _text(case.get("type")),
        "group": _text(case.get("group")),
        "polarity": _text(case.get("polarity")),
        "title": _text(case.get("title")),
        "design_method": _text(case.get("design_method")),
        "priority": _text(case.get("priority")),
        "pre_operations": normalize_list_items(case.get("pre_operations")),
        "test_data": normalize_list_items(case.get("test_data")),
        "steps": normalize_list_items(case.get("steps")),
        "expected_results": normalize_list_items(case.get("expected_results")),
        "post_operations": normalize_list_items(case.get("post_operations")),
        "traceability_refs": normalized_refs,
    }


def normalize_baseline(payload: Any) -> dict[str, Any]:
    if not isinstance(payload, dict):
        raise ValueError("IR 根节点必须是对象。")
    raw_entries = payload.get("traceability")
    if isinstance(raw_entries, (list, tuple)):
        entries = [dict(entry) if isinstance(entry, dict) else entry for entry in raw_entries]
    elif isinstance(raw_entries, dict):
        entries = [dict(entry) if isinstance(entry, dict) else entry for entry in raw_entries.values()]
    else:
        entries = []
    raw_cases = payload.get("cases")
    cases = [normalize_case(case) for case in raw_cases] if isinstance(raw_cases, (list, tuple)) else []
    return {
        "schema_version": SCHEMA_VERSION,
        "meta": normalize_meta(payload.get("meta")),
        "traceability": entries,
        "cases": cases,
    }


def _read_json(path: str | Path) -> Any:
    source = Path(path)
    try:
        text = source.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as exc:
        raise ValueError(f"无法读取 IR 文件: {source}（{exc}）") from None
    try:
        return json.loads(text)
    except json.JSONDecodeError as exc:
        raise ValueError(f"IR 文件 JSON 格式错误: {source}（{exc}）") from None


def load_baseline(path: str | Path) -> dict[str, Any]:
    return normalize_baseline(_read_json(path))


def load_baseline_raw(path: str | Path) -> Any:
    """Read the IR JSON without normalizing, so validation can still see rejected fields."""
    return _read_json(path)


def validate_baseline(
    payload: Any,
    require_design_method: bool = False,
    allow_unassigned_ids: bool = False,
    require_finalized_traceability: bool = True,
    require_tree_fields: bool = True,
    check_traceability_links: bool = True,
    allow_unsettled_coverage: bool = False,
) -> list[str]:
    errors: list[str] = []
    if not isinstance(payload, dict):
        return ["IR 根节点必须是对象。"]

    raw_meta = payload.get("meta")
    if isinstance(raw_meta, dict) and "case_count" in raw_meta:
        errors.append("meta.case_count 不得填写，用例数量按 cases 长度自动计算。")

    ir = normalize_baseline(payload)
    meta = ir["meta"]

    if not meta["requirement_name"]:
        errors.append("meta.requirement_name 不能为空。")
    if _has_newline(meta["requirement_name"]):
        errors.append("meta.requirement_name 不得包含换行。")
    if not is_valid_date_or_placeholder(meta["release_date"]):
        errors.append("meta.release_date 必须为有效 YYYY-MM-DD 或 未指定日期。")
    if not meta["title"]:
        errors.append("meta.title 不能为空。")
    elif _has_newline(meta["title"]):
        errors.append("meta.title 不得包含换行。")
    if meta["sheet_name"] in RESERVED_SHEET_NAMES:
        errors.append(
            f"meta.sheet_name 不得使用保留名: {'、'.join(RESERVED_SHEET_NAMES)}。"
        )
    if _contains_marker(payload.get("meta")):
        errors.append("meta 不得保留 [待确认] 标记，请改为待确认条目。")
    if _contains_marker(payload.get("traceability")):
        errors.append("traceability 不得保留 [待确认] 标记，请使用 status=待确认 条目。")

    entries, entry_errors = validate_traceability_entries(
        ir["traceability"],
        allow_unassigned_ids=allow_unassigned_ids,
        allow_unsettled_links=allow_unsettled_coverage,
    )
    errors.extend(entry_errors)
    # IDs that were declared but rejected for other reasons: used only to avoid
    # cascading "引用不存在的追溯 ID" errors, never to satisfy link checks.
    declared_statuses: dict[str, str] = {}
    for raw_entry in ir["traceability"]:
        if not isinstance(raw_entry, dict):
            continue
        raw_id = _text(raw_entry.get("id"))
        if raw_id and raw_id not in declared_statuses:
            declared_statuses[raw_id] = _text(raw_entry.get("status"))
    if require_finalized_traceability:
        draft = [entry["id"] for entry in entries.values() if entry["status"] == "待覆盖"]
        if draft:
            errors.append(
                "以下追溯条目仍为待覆盖，请补齐用例或移除条目后再渲染: "
                + "、".join(draft)
                + "。"
            )

    cases = ir["cases"]
    if not cases:
        errors.append("cases 必须为非空数组。")

    raw_cases = payload.get("cases")
    if isinstance(raw_cases, (list, tuple)):
        for index, raw_case in enumerate(raw_cases, start=1):
            if not isinstance(raw_case, dict):
                continue
            for field, hint in (
                ("requirement_id", "请改用 traceability_refs"),
                ("notes", "用例不支持备注字段，请删除"),
                ("preconditions", "请改用 pre_operations"),
                ("post_cleanup", "请改用 post_operations"),
            ):
                if field in raw_case:
                    errors.append(
                        f"第 {index} 条用例包含不支持字段 {field}（{hint}）。"
                    )
            raw_refs = raw_case.get("traceability_refs")
            if isinstance(raw_refs, (list, tuple)) and any(
                item is None or not str(item).strip() for item in raw_refs
            ):
                errors.append(f"第 {index} 条用例的 traceability_refs 存在空项。")

    seen_ids: set[str] = set()
    seen_titles: set[str] = set()
    trace_order = [
        trace_id for trace_id, entry in entries.items() if entry["status"] != "待确认"
    ]
    for index, case in enumerate(cases, start=1):
        label = f"第 {index} 条用例"
        case_id = case["id"]
        if case_id in PLACEHOLDER_IDS:
            if not allow_unassigned_ids:
                errors.append(
                    f"{label}缺少用例 ID，请先执行 --assign-ids 或在 IR 中填写 TC-{{需求简写}}-流水号。"
                )
        else:
            if ID_INVALID_PATTERN.search(case_id):
                errors.append(f"{label}的用例 ID '{case_id}' 含非法字符（空白、逗号或方括号）。")
            elif case_id in seen_ids:
                errors.append(f"用例 ID 重复: '{case_id}'。")
            seen_ids.add(case_id)

        if require_tree_fields:
            if not case["module"]:
                errors.append(f"{label}（{case_id or '无 ID'}）的 module 不能为空。")
            elif _has_newline(case["module"]):
                errors.append(f"{label}（{case_id}）的 module 不得包含换行。")
        if case["type"] not in ALLOWED_CASE_TYPES:
            errors.append(
                f"{label}（{case_id}）的 type 必须为 {'、'.join(ALLOWED_CASE_TYPES)} 之一。"
            )
        if require_tree_fields and case["polarity"] not in ALLOWED_POLARITIES:
            errors.append(
                f"{label}（{case_id}）的 polarity 必须为 {' / '.join(ALLOWED_POLARITIES)}。"
            )
        if case["priority"] not in ALLOWED_PRIORITIES:
            errors.append(
                f"{label}（{case_id}）的 priority 必须为 {'、'.join(ALLOWED_PRIORITIES)} 之一。"
            )
        if case["group"] and _has_newline(case["group"]):
            errors.append(f"{label}（{case_id}）的 group 不得包含换行。")

        title = case["title"]
        if not title:
            errors.append(f"{label}（{case_id}）的 title 不能为空。")
        else:
            if _has_newline(title):
                errors.append(f"{label}（{case_id}）的 title 不得包含换行。")
            elif title.startswith(("正向-", "反向-")):
                errors.append(f"{label}（{case_id}）的 title 不得以“正向-”或“反向-”开头。")
            elif re.search(r"-P[0-4]$", title):
                errors.append(f"{label}（{case_id}）的 title 不得以“-P0..P4”结尾。")
            elif title in seen_titles:
                errors.append(f"用例标题重复: '{title}'。")
            seen_titles.add(title)

        for field in REQUIRED_FIELDS:
            if not case[FIELD_KEY_MAP[field]]:
                errors.append(f"{label}（{case_id}）缺少必填字段“{field}”的内容。")

        if require_design_method:
            if not case["design_method"]:
                errors.append(f"{label}（{case_id}）缺少 design_method，生成 Excel 时必须填写。")
            elif case["design_method"] not in ALLOWED_DESIGN_METHODS:
                errors.append(
                    f"{label}（{case_id}）的 design_method 必须为 "
                    f"{'、'.join(ALLOWED_DESIGN_METHODS)} 之一。"
                )
        elif case["design_method"] and case["design_method"] not in ALLOWED_DESIGN_METHODS:
            errors.append(
                f"{label}（{case_id}）的 design_method 必须为 "
                f"{'、'.join(ALLOWED_DESIGN_METHODS)} 之一。"
            )

        refs = case["traceability_refs"]
        if not isinstance(refs, list):
            errors.append(f"{label}（{case_id}）的 traceability_refs 必须为数组。")
            refs = []
        if not refs:
            errors.append(f"{label}（{case_id}）的 traceability_refs 不能为空。")
        elif len(set(refs)) != len(refs):
            errors.append(f"{label}（{case_id}）的 traceability_refs 存在重复项。")
        for ref in refs:
            entry = entries.get(ref)
            if entry is not None:
                if entry["status"] == "待确认":
                    errors.append(f"{label}（{case_id}）不得引用待确认条目 '{ref}'。")
            elif ref in declared_statuses:
                # The entry exists but was rejected (missing rule/source, bad
                # case_ids, duplicate ID...). Its own error is already reported,
                # so do not add a misleading "missing trace ID" cascade.
                if declared_statuses[ref] == "待确认":
                    errors.append(f"{label}（{case_id}）不得引用待确认条目 '{ref}'。")
            else:
                errors.append(f"{label}（{case_id}）引用了不存在的追溯 ID '{ref}'。")
        if (
            check_traceability_links
            and refs
            and all(ref in entries for ref in refs)
        ):
            expected_refs = [trace_id for trace_id in trace_order if trace_id in refs]
            diff = describe_traceability_mismatch(expected_refs, refs)
            if diff:
                errors.append(
                    f"{label}（{case_id}）的追溯引用顺序与追溯表不一致：{diff}"
                )

        for field, value in case.items():
            if _contains_marker(value):
                errors.append(f"{label}（{case_id}）的 {field} 不得保留 [待确认] 标记。")

    if not check_traceability_links:
        return errors

    ordered_cases = canonical_tree_order(cases) if require_tree_fields else cases
    expected_links: dict[str, list[str]] = {trace_id: [] for trace_id in entries}
    for case in ordered_cases:
        refs = case["traceability_refs"]
        if not isinstance(refs, list):
            continue
        if allow_unassigned_ids and case["id"] in PLACEHOLDER_IDS:
            continue
        for ref in refs:
            if ref in expected_links and case["id"] and case["id"] not in expected_links[ref]:
                expected_links[ref].append(case["id"])
    for trace_id, entry in entries.items():
        if entry["status"] == "待确认":
            continue
        if not require_finalized_traceability and entry["status"] == "待覆盖":
            continue
        diff = describe_traceability_mismatch(entry["case_ids"], expected_links[trace_id])
        if diff:
            errors.append(f"追溯条目 '{trace_id}' 的关联用例与正文引用不一致：{diff}")

    return errors


CASE_COMPARE_FIELDS = (
    "id",
    "title",
    "type",
    "priority",
    "design_method",
    "pre_operations",
    "test_data",
    "steps",
    "expected_results",
    "post_operations",
    "traceability_refs",
)
TREE_COMPARE_FIELDS = ("module", "group", "polarity")
TRACE_COMPARE_FIELDS = ("source", "location", "rule", "issue", "assumption", "question")


def _short(value: Any, limit: int = 40) -> str:
    text = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False)
    return text if len(text) <= limit else text[:limit] + "…"


def _trace_signature(entry: dict[str, Any]) -> tuple[str, ...]:
    signature = [_text(entry.get(field)) for field in TRACE_COMPARE_FIELDS]
    signature.append("status=" + _text(entry.get("status")))
    case_ids = ",".join(_text(item) for item in entry.get("case_ids") or [] if _text(item))
    signature.append("case_ids=" + case_ids)
    return tuple(signature)


def describe_ir_differences(
    base: Any,
    other: Any,
    base_label: str = "Markdown",
    other_label: str = "Excel",
    include_tree_fields: bool = True,
    include_design_method: bool = True,
) -> list[str]:
    """Compare two IRs and describe every case and traceability difference.

    ``include_design_method=False`` is used when one side is Markdown, which
    carries no design method, so the absence is not a real difference.
    """
    base_ir = normalize_baseline(base)
    other_ir = normalize_baseline(other)
    errors: list[str] = []

    base_ids = [case["id"] for case in base_ir["cases"]]
    other_ids = [case["id"] for case in other_ir["cases"]]
    if set(base_ids) != set(other_ids):
        missing = [case_id for case_id in base_ids if case_id not in other_ids]
        extra = [case_id for case_id in other_ids if case_id not in base_ids]
        if missing:
            errors.append(f"{other_label} 缺少用例: {', '.join(missing)}。")
        if extra:
            errors.append(f"{other_label} 多出用例: {', '.join(extra)}。")
    elif base_ids != other_ids:
        errors.append(
            f"用例顺序不一致：{base_label}=[{', '.join(base_ids)}]，"
            f"{other_label}=[{', '.join(other_ids)}]。"
        )

    fields = CASE_COMPARE_FIELDS + (TREE_COMPARE_FIELDS if include_tree_fields else ())
    if not include_design_method:
        fields = tuple(field for field in fields if field != "design_method")
    other_cases = {case["id"]: case for case in other_ir["cases"]}
    for case in base_ir["cases"]:
        counterpart = other_cases.get(case["id"])
        if counterpart is None:
            continue
        for field in fields:
            if case[field] != counterpart[field]:
                errors.append(
                    f"用例 {case['id']} 的 {field} 不一致："
                    f"{base_label}={_short(case[field])}，{other_label}={_short(counterpart[field])}。"
                )

    base_entries = {
        _text(entry.get("id")): entry
        for entry in base_ir["traceability"]
        if isinstance(entry, dict) and _text(entry.get("id"))
    }
    other_entries = {
        _text(entry.get("id")): entry
        for entry in other_ir["traceability"]
        if isinstance(entry, dict) and _text(entry.get("id"))
    }
    for trace_id in base_entries:
        if trace_id not in other_entries:
            errors.append(f"{other_label} 缺少追溯条目: {trace_id}。")
    for trace_id in other_entries:
        if trace_id not in base_entries:
            errors.append(f"{other_label} 多出追溯条目: {trace_id}。")
    for trace_id, entry in base_entries.items():
        counterpart = other_entries.get(trace_id)
        if counterpart is None:
            continue
        expected = _trace_signature(entry)
        actual = _trace_signature(counterpart)
        if expected != actual:
            differences = [
                f"{label}（{base_label}={_short(left)}，{other_label}={_short(right)}）"
                for label, left, right in zip(
                    (*TRACE_COMPARE_FIELDS, "status", "case_ids"), expected, actual
                )
                if left != right
            ]
            errors.append(f"追溯条目 {trace_id} 不一致：{'；'.join(differences)}。")

    base_order = [entry for entry in base_entries]
    other_order = [entry for entry in other_entries]
    if set(base_order) == set(other_order) and base_order != other_order:
        errors.append(
            f"追溯条目顺序不一致：{base_label}=[{', '.join(base_order)}]，"
            f"{other_label}=[{', '.join(other_order)}]。"
        )
    return errors


def _most_common_prefix(cases: list[dict[str, Any]]) -> str:
    counts: Counter[str] = Counter()
    order: list[str] = []
    for case in cases:
        match = ID_PATTERN.match(case["id"])
        if not match:
            continue
        prefix = match.group(1)
        if prefix not in counts:
            order.append(prefix)
        counts[prefix] += 1
    if not counts:
        return ""
    return max(order, key=lambda prefix: (counts[prefix], -order.index(prefix)))


def assign_missing_ids(cases: list[dict[str, Any]], id_prefix: str | None = None) -> list[str]:
    used = {case["id"] for case in cases if case["id"] not in PLACEHOLDER_IDS}
    prefix_max: dict[str, int] = {}
    for case_id in used:
        match = ID_PATTERN.match(case_id)
        if match:
            prefix = match.group(1)
            prefix_max[prefix] = max(prefix_max.get(prefix, 0), int(match.group(2)))

    prefix = _text(id_prefix) or _most_common_prefix(cases) or "TC-CASE"
    number = prefix_max.get(prefix, 0)
    assigned: list[str] = []
    for case in cases:
        if case["id"] not in PLACEHOLDER_IDS:
            continue
        number += 1
        candidate = f"{prefix}-{number:03d}"
        while candidate in used:
            number += 1
            candidate = f"{prefix}-{number:03d}"
        case["id"] = candidate
        used.add(candidate)
        assigned.append(candidate)
    return assigned


def canonical_tree_order(cases: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Flatten cases in the rendered tree order: module → type → group, each by first appearance."""
    tree: dict[str, dict[str, dict[str, list[dict[str, Any]]]]] = {}
    for case in cases:
        tree.setdefault(case.get("module", ""), {}).setdefault(
            case.get("type", ""), {}
        ).setdefault(case.get("group", ""), []).append(case)
    return [
        case
        for types in tree.values()
        for groups in types.values()
        for bucket in groups.values()
        for case in bucket
    ]


def prepare_render_ir(
    raw: Any,
    assign_ids: bool = False,
    id_prefix: str | None = None,
    require_design_method: bool = False,
) -> tuple[dict[str, Any], list[str], list[str]]:
    """Validate the draft IR, settle IDs and traceability, then validate again.

    This is the single owner of the render pipeline: draft pre-check →
    normalization/settlement → final full validation.
    """
    errors = validate_baseline(
        raw,
        allow_unassigned_ids=assign_ids,
        require_finalized_traceability=False,
        check_traceability_links=False,
        allow_unsettled_coverage=True,
    )
    if errors:
        return {}, errors, []

    ir = normalize_baseline(raw)
    if assign_ids:
        assign_missing_ids(ir["cases"], id_prefix)
    ir["cases"] = canonical_tree_order(ir["cases"])

    entries, sync_errors, warnings = sync_traceability_from_cases(
        ir["traceability"], ir["cases"]
    )
    if sync_errors:
        return {}, sync_errors, warnings
    ir["traceability"] = list(entries.values())

    errors = validate_baseline(ir, require_design_method=require_design_method)
    if errors:
        return {}, errors, warnings
    return ir, [], warnings


def renumber_cases(
    cases: list[dict[str, Any]],
    traceability: Any,
    id_prefix: str | None = None,
) -> dict[str, str]:
    ordered = canonical_tree_order(cases)
    cases[:] = ordered
    default_prefix = _most_common_prefix(cases) or "TC-CASE"
    explicit = _text(id_prefix)
    counters: dict[str, int] = {}
    mapping: dict[str, str] = {}
    for case in cases:
        old_id = case["id"]
        prefix = explicit or (ID_PATTERN.match(old_id).group(1) if ID_PATTERN.match(old_id) else default_prefix)
        counters[prefix] = counters.get(prefix, 0) + 1
        new_id = f"{prefix}-{counters[prefix]:03d}"
        mapping[old_id] = new_id
        case["id"] = new_id

    order = {case["id"]: index for index, case in enumerate(cases)}
    entries = traceability.values() if isinstance(traceability, dict) else traceability or []
    for entry in entries:
        if not isinstance(entry, dict) or entry.get("status") == "待确认":
            continue
        mapped = [mapping.get(case_id, case_id) for case_id in entry.get("case_ids") or []]
        entry["case_ids"] = sorted(
            dict.fromkeys(mapped), key=lambda case_id: order.get(case_id, len(order))
        )
    return mapping


# ---------------------------------------------------------------------------
# Markdown parsing
# ---------------------------------------------------------------------------


def _indent_of(line: str) -> int:
    return len(line) - len(line.lstrip(" "))


def _parse_fields(
    lines: list[str],
    start_line_no: int,
    case_indent: int,
    errors: list[str],
    warnings: list[str],
) -> dict[str, list[str]]:
    fields: dict[str, list[str]] = {}
    order: list[str] = []
    current: str | None = None
    field_indent = case_indent + 2

    for offset, raw_line in enumerate(lines):
        if not raw_line.strip():
            continue
        line_no = start_line_no + offset
        indent = _indent_of(raw_line)
        stripped = raw_line.strip()

        if indent == field_indent:
            match = FIELD_PATTERN.match(stripped)
            if match:
                name = match.group(1)
                if name in fields:
                    errors.append(f"第 {line_no} 行：字段“{name}”重复出现。")
                    current = None
                    continue
                index = FIELD_ORDER.index(name)
                if order and index <= FIELD_ORDER.index(order[-1]):
                    errors.append(
                        f"第 {line_no} 行：字段“{name}”顺序错误，"
                        f"必须依次为 {' → '.join(FIELD_ORDER)}。"
                    )
                fields[name] = []
                order.append(name)
                current = name
                continue
            if stripped.startswith("- "):
                name = stripped[2:].strip()
                if name in UNSUPPORTED_FIELDS:
                    errors.append(
                        f"第 {line_no} 行：不支持字段“{name}”，"
                        "请改用 前置操作/测试步骤/预期结果/后置操作。"
                    )
                else:
                    errors.append(f"第 {line_no} 行：未知的直接子字段“{name}”。")
                current = None
                continue
            errors.append(f"第 {line_no} 行：用例字段必须使用“- 字段名”格式。")
            current = None
            continue

        if indent < field_indent:
            errors.append(f"第 {line_no} 行：用例字段缩进不足（应为 {field_indent} 个空格）。")
            continue

        if current is None:
            errors.append(f"第 {line_no} 行：列表项之前缺少字段标题。")
            continue

        item = ITEM_PATTERN.match(stripped)
        if item and indent == field_indent + 2:
            fields[current].append(item.group(1).strip())
            continue

        continuation = re.sub(r"^[>\-*]\s*", "", stripped).strip()
        if not continuation:
            continue
        if fields[current]:
            fields[current][-1] = fields[current][-1] + "；" + continuation
            warnings.append(
                f"第 {line_no} 行：缩进续行已合并到上一条列表项：{continuation}"
            )
        else:
            errors.append(
                f"第 {line_no} 行：字段“{current}”的首个列表项之前出现续行，无法合并：{continuation}"
            )

    return fields


def _parse_covered_table(
    rows: list[list[str]], errors: list[str]
) -> list[dict[str, Any]]:
    entries: list[dict[str, Any]] = []
    for row_index, row in enumerate(rows):
        if not any(cell.strip() for cell in row):
            continue
        if len(row) < len(COVERED_HEADERS):
            errors.append(f"需求追溯表第 {row_index + 1} 行列数不足，应为 5 列。")
            continue
        trace_id, source, location, rule, case_text = row[:5]
        case_ids, case_errors = parse_trace_id_list(case_text) if case_text.strip() else ([], [])
        errors.extend(f"需求追溯表“{trace_id}”：{message}" for message in case_errors)
        entries.append(
            {
                "id": trace_id.strip(),
                "source": source.strip(),
                "location": location.strip(),
                "rule": rule.strip(),
                "status": "已覆盖",
                "case_ids": case_ids,
            }
        )
    return entries


def _parse_pending_table(rows: list[list[str]], errors: list[str]) -> list[dict[str, Any]]:
    entries: list[dict[str, Any]] = []
    for row_index, row in enumerate(rows):
        if not any(cell.strip() for cell in row):
            continue
        if len(row) < len(PENDING_HEADERS):
            errors.append(f"待确认项第 {row_index + 1} 行列数不足，应为 6 列。")
            continue
        trace_id, source, location, issue, assumption, question = row[:6]
        entries.append(
            {
                "id": trace_id.strip(),
                "source": source.strip(),
                "location": location.strip(),
                "status": "待确认",
                "issue": issue.strip(),
                "assumption": assumption.strip(),
                "question": question.strip(),
            }
        )
    return entries


def _parse_traceability_section(
    lines: list[str], errors: list[str], require_covered_table: bool = True
) -> list[dict[str, Any]]:
    entries: list[dict[str, Any]] = []
    section: str | None = None
    awaiting_table: str | None = None
    saw_covered_table = False
    index = 0
    while index < len(lines):
        raw = lines[index]
        if not raw.strip():
            index += 1
            continue
        if raw.startswith("### "):
            name = raw[4:].strip()
            if awaiting_table is not None:
                errors.append(f"“### {awaiting_table}”标题后缺少表格。")
                awaiting_table = None
            if name == "需求追溯表":
                section = "covered"
                awaiting_table = name
            elif name == "待确认项":
                section = "pending"
                awaiting_table = name
            else:
                errors.append(f"需求追溯小节出现未知三级标题“{name}”。")
                section = None
            index += 1
            continue
        if raw.startswith("#"):
            errors.append("需求追溯小节不得出现二级及以上标题。")
            index += 1
            continue
        if not raw.strip().startswith("|"):
            errors.append(f"需求追溯小节存在表外内容：{raw.strip()[:30]}")
            index += 1
            continue

        table_lines: list[str] = []
        while index < len(lines) and lines[index].strip().startswith("|"):
            table_lines.append(lines[index])
            index += 1
        awaiting_table = None
        if section is None:
            errors.append("需求追溯表格之前缺少 ### 需求追溯表 或 ### 待确认项 标题。")
            continue
        if section == "covered":
            saw_covered_table = True

        rows = [split_md_table_row(line) for line in table_lines]
        expected = COVERED_HEADERS if section == "covered" else PENDING_HEADERS
        if len(rows) < 2:
            errors.append("需求追溯表格必须包含表头与分隔行。")
            continue
        if tuple(rows[0]) != expected:
            errors.append(f"需求追溯表头必须为：{' | '.join(expected)}。")
            continue
        if len(rows[1]) != len(expected) or not all(
            TOKEN_PATTERN.match(cell) for cell in rows[1]
        ):
            errors.append("需求追溯表格缺少标准分隔行。")
            continue
        before = len(entries)
        if section == "covered":
            entries.extend(_parse_covered_table(rows[2:], errors))
        else:
            entries.extend(_parse_pending_table(rows[2:], errors))
        if len(entries) == before:
            label = "需求追溯表" if section == "covered" else "待确认项"
            errors.append(f"{label}表格不得为空。")
    if awaiting_table is not None:
        errors.append(f"“### {awaiting_table}”标题后缺少表格。")
    if require_covered_table and not saw_covered_table:
        errors.append("需求追溯小节缺少 ### 需求追溯表 表格。")
    return entries


def parse_markdown_text(
    text: str,
    allow_unassigned_ids: bool = False,
    strict_count: bool = True,
) -> tuple[dict[str, Any], list[str], list[str]]:
    """Parse Markdown document text into the merged IR."""
    lines = text.splitlines()
    separators = [i for i, line in enumerate(lines) if re.match(r"^##\s+需求追溯\s*$", line)]
    errors: list[str] = []
    warnings: list[str] = []
    if not separators:
        errors.append("缺少唯一的 ## 需求追溯 小节。")
        tree_lines = lines
        trace_lines: list[str] = []
    elif len(separators) > 1:
        errors.append("文档中出现多个 ## 需求追溯 小节，必须唯一。")
        tree_lines = lines[: separators[0]]
        trace_lines = lines[separators[0] + 1 :]
    else:
        tree_lines = lines[: separators[0]]
        trace_lines = lines[separators[0] + 1 :]

    title = ""
    for raw in tree_lines:
        match = re.match(r"^#\s+(.+?)\s*$", raw)
        if match:
            title = match.group(1)
            break
    if not title:
        errors.append("文档首行必须为一级标题“# 测试用例 - {日期} - [{需求名称}]”。")
    meta = normalize_meta({})
    title_match = TITLE_PATTERN.match(title) if title else None
    if title_match:
        meta["release_date"] = title_match.group(1).strip()
        meta["requirement_name"] = title_match.group(2).strip()
        meta["title"] = title
    elif title:
        errors.append("一级标题必须为“测试用例 - {日期} - [{需求名称}]”格式。")
        meta["title"] = title
        meta["requirement_name"] = title

    declared_counts: list[int] = []
    for raw in tree_lines:
        match = COUNT_PATTERN.match(raw.strip())
        if match:
            declared_counts.append(int(match.group(1)))

    cases: list[dict[str, Any]] = []
    module = ""
    type_name = ""
    group = ""
    module_nodes: list[tuple[str, int]] = []
    type_nodes: list[tuple[str, str, int]] = []
    group_nodes: list[tuple[str, str, str, int]] = []
    used_modules: set[str] = set()
    used_types: set[tuple[str, str]] = set()
    used_groups: set[tuple[str, str, str]] = set()
    title_line_no: int | None = None
    index = 0
    while index < len(tree_lines):
        raw = tree_lines[index]
        if not raw.strip():
            index += 1
            continue
        stripped = raw.strip()
        line_no = index + 1
        if stripped.startswith("#"):
            is_title = stripped.startswith("# ") and stripped[2:].strip() == title
            if stripped.startswith("####"):
                errors.append(f"第 {line_no} 行：不得使用四级及更深标题。")
            elif is_title and title_line_no is None:
                title_line_no = line_no
            else:
                errors.append(
                    f"第 {line_no} 行：任务树中不得出现标题“{stripped[:30]}”。"
                )
            index += 1
            continue
        if COUNT_PATTERN.match(stripped):
            index += 1
            continue
        if re.match(r"^- \[x\]", stripped, re.IGNORECASE):
            errors.append(f"第 {line_no} 行：不得使用已完成标记 - [x]。")
            index += 1
            continue

        case_match = CASE_LINE_PATTERN.match(raw)
        if case_match:
            case_indent = len(case_match.group(1))
            polarity = case_match.group(2)
            case_title = case_match.group(3)
            priority = case_match.group(4)
            if case_indent not in (4, 6):
                errors.append(f"第 {line_no} 行：用例缩进必须为 4 或 6 个空格。")
            if not module or not type_name:
                errors.append(f"第 {line_no} 行：用例之前缺少功能模块或一级测试类型节点。")
            case_group = group if case_indent == 6 else ""
            title_hit = CASE_ID_IN_TITLE_PATTERN.search(case_title)
            if title_hit:
                errors.append(
                    f"第 {line_no} 行：用例标题不得包含用例 ID '{title_hit.group(0)}'，"
                    "请移至下方引用行。"
                )

            end = index + 1
            while end < len(tree_lines):
                candidate = tree_lines[end]
                if not candidate.strip():
                    end += 1
                    continue
                if _indent_of(candidate) > case_indent:
                    end += 1
                    continue
                break
            block = tree_lines[index + 1 : end]

            first_offset = 0
            while first_offset < len(block) and not block[first_offset].strip():
                first_offset += 1
            ref_match = (
                REF_LINE_PATTERN.match(block[first_offset])
                if first_offset < len(block)
                else None
            )
            case_id = ""
            refs: list[str] = []
            if ref_match is None:
                errors.append(
                    f"第 {line_no} 行：用例标题下首个非空行必须是引用行“> [TC-ID] 追溯: TRACE-ID”。"
                )
            else:
                if _indent_of(block[first_offset]) != case_indent + 2:
                    errors.append(
                        f"第 {line_no + first_offset + 1} 行：引用行缩进必须为 "
                        f"{case_indent + 2} 个空格。"
                    )
                case_id = (ref_match.group(1) or "").strip()
                refs, ref_errors = parse_trace_id_list(ref_match.group(2))
                errors.extend(f"第 {line_no} 行：{message}" for message in ref_errors)
                if not case_id:
                    if not allow_unassigned_ids:
                        errors.append(
                            f"第 {line_no} 行：引用行缺少用例 ID（如需先补号请使用 --assign-ids）。"
                        )
                elif case_id in ("TC-NEW", "NEW") and not allow_unassigned_ids:
                    errors.append(
                        f"第 {line_no} 行：引用行存在占位用例 ID '{case_id}'，请先执行 --assign-ids。"
                    )

            body_start = first_offset + 1 if ref_match is not None else first_offset
            fields = _parse_fields(
                block[body_start:],
                line_no + 1 + body_start,
                case_indent,
                errors,
                warnings,
            )
            for name in REQUIRED_FIELDS:
                if name not in fields:
                    errors.append(f"第 {line_no} 行：用例缺少必填字段“{name}”。")
                elif not fields[name]:
                    errors.append(f"第 {line_no} 行：字段“{name}”下缺少列表项。")
            for name in ("测试数据", "后置操作"):
                if name in fields and not fields[name]:
                    errors.append(f"第 {line_no} 行：字段“{name}”下缺少列表项。")

            cases.append(
                {
                    "id": case_id,
                    "module": module,
                    "type": type_name,
                    "group": case_group,
                    "polarity": polarity,
                    "title": case_title.strip(),
                    "design_method": "",
                    "priority": priority,
                    "pre_operations": fields.get("前置操作", []),
                    "test_data": fields.get("测试数据", []),
                    "steps": fields.get("测试步骤", []),
                    "expected_results": fields.get("预期结果", []),
                    "post_operations": fields.get("后置操作", []),
                    "traceability_refs": refs,
                }
            )
            used_modules.add(module)
            used_types.add((module, type_name))
            if case_group:
                used_groups.add((module, type_name, case_group))
            index = end
            continue

        node = NODE_PATTERN.match(stripped)
        if node:
            name = node.group(1).strip()
            indent = _indent_of(raw)
            if indent == 0:
                module = name
                type_name = ""
                group = ""
                module_nodes.append((name, line_no))
            elif indent == 2:
                if not module:
                    errors.append(f"第 {line_no} 行：一级测试类型之前缺少功能模块节点。")
                type_name = name
                group = ""
                type_nodes.append((module, name, line_no))
            elif indent == 4:
                if not type_name:
                    errors.append(f"第 {line_no} 行：分组之前缺少一级测试类型节点。")
                group = name
                group_nodes.append((module, type_name, name, line_no))
            else:
                errors.append(f"第 {line_no} 行：任务树层级必须为 0/2/4 个空格。")
            index += 1
            continue

        errors.append(f"第 {line_no} 行：无法识别的任务树内容：{stripped[:30]}")
        index += 1

    for name, node_line in module_nodes:
        if name not in used_modules:
            errors.append(f"第 {node_line} 行：功能模块“{name}”下没有任何用例。")
    for module_name, type_value, node_line in type_nodes:
        if (module_name, type_value) not in used_types:
            errors.append(f"第 {node_line} 行：测试类型“{type_value}”下没有任何用例。")
    for module_name, type_value, group_name, node_line in group_nodes:
        if (module_name, type_value, group_name) not in used_groups:
            errors.append(f"第 {node_line} 行：分组“{group_name}”下没有任何用例。")

    if not declared_counts:
        message = "缺少“> 用例数量: 用例 N 条”声明。"
        (errors if strict_count else warnings).append(message)
    elif len(declared_counts) > 1:
        message = f"用例数量声明重复出现 {len(declared_counts)} 次。"
        (errors if strict_count else warnings).append(message)
    elif declared_counts[0] != len(cases):
        message = f"用例数量声明为 {declared_counts[0]} 条，实际解析到 {len(cases)} 条。"
        (errors if strict_count else warnings).append(message)

    entries = _parse_traceability_section(
        trace_lines, errors, require_covered_table=len(separators) == 1
    )
    if any(not case["design_method"] for case in cases):
        warnings.append(
            "Markdown 不包含设计方法，导出 IR 后需从 Excel 合并或手工补充 design_method。"
        )

    ir = {
        "schema_version": SCHEMA_VERSION,
        "meta": meta,
        "traceability": entries,
        "cases": cases,
    }
    return ir, errors, warnings


def parse_markdown_to_ir(
    md_path: str | Path,
    allow_unassigned_ids: bool = False,
    strict_count: bool = True,
) -> tuple[dict[str, Any], list[str], list[str]]:
    path = Path(md_path)
    try:
        text = path.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as exc:
        return {}, [f"无法读取 Markdown 文件: {path}（{exc}）"], []
    return parse_markdown_text(text, allow_unassigned_ids, strict_count)


# ---------------------------------------------------------------------------
# Excel parsing
# ---------------------------------------------------------------------------


def _cell_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _require_openpyxl() -> Any:
    try:
        import openpyxl
    except ImportError:
        return None
    return openpyxl


def _sheet_headers(ws: Any, row: int, count: int) -> tuple[str, ...]:
    return tuple(_cell_text(ws.cell(row=row, column=column).value) for column in range(1, count + 1))


def _module_from_title(title: str) -> str:
    match = re.match(r"^\[([^\]]+)\]", title)
    return match.group(1).strip() if match else ""


def _parse_info_sheet(workbook: Any, errors: list[str]) -> dict[str, Any]:
    if "生成说明" not in workbook.sheetnames:
        errors.append("工作簿缺少“生成说明”工作表。")
        return {}
    ws = workbook["生成说明"]
    data: dict[str, str] = {}
    for row in range(2, ws.max_row + 1):
        label = _cell_text(ws.cell(row=row, column=1).value)
        if not label:
            continue
        value_cell = ws.cell(row=row, column=2)
        if value_cell.data_type == "f":
            errors.append(f"生成说明第 {row} 行包含公式，必须为纯文本。")
            continue
        value = value_cell.value
        data[label] = "" if value is None else str(value)
    return {
        "requirement_name": data.get("需求名称", ""),
        "release_date": data.get("上线日期", ""),
        "coverage": data.get("覆盖范围", ""),
        "notes": data.get("补充说明", ""),
    }


def _parse_v2_case_sheet(
    ws: Any, errors: list[str], warnings: list[str]
) -> list[dict[str, Any]]:
    headers = _sheet_headers(ws, 2, len(V2_HEADERS))
    if headers != V2_HEADERS:
        errors.append(f"主工作表第 2 行表头必须为：{' | '.join(V2_HEADERS)}。")
        return []

    cases: list[dict[str, Any]] = []
    for row in range(3, ws.max_row + 1):
        cells = [
            ws.cell(row=row, column=column) for column in range(1, len(V2_HEADERS) + 1)
        ]
        values = [cell.value for cell in cells]
        if all(value is None or str(value).strip() == "" for value in values):
            continue
        for column, cell in enumerate(cells, start=1):
            if cell.data_type == "f":
                errors.append(f"主工作表第 {row} 行第 {column} 列包含公式，必须为纯文本。")
        trace_ids, trace_errors = (
            parse_trace_id_list(_cell_text(values[0])) if _cell_text(values[0]) else ([], [])
        )
        errors.extend(f"主工作表第 {row} 行：{message}" for message in trace_errors)
        title = _cell_text(values[2])
        cases.append(
            {
                "id": _cell_text(values[1]),
                "module": _module_from_title(title),
                "type": _cell_text(values[3]),
                "group": "",
                "polarity": "",
                "title": title,
                "design_method": _cell_text(values[4]),
                "priority": _cell_text(values[5]),
                "pre_operations": normalize_list_items(values[6]),
                "test_data": normalize_list_items(values[7]),
                "steps": normalize_list_items(values[8]),
                "expected_results": normalize_list_items(values[9]),
                "post_operations": normalize_list_items(values[10]),
                "traceability_refs": trace_ids,
            }
        )
    if any(not case["polarity"] for case in cases):
        warnings.append("Excel 不包含正向/反向信息，请在导出的 IR 中补充 polarity。")
    return cases


def _parse_traceability_sheet(
    ws: Any, status: str, errors: list[str]
) -> list[dict[str, Any]]:
    expected = COVERED_HEADERS_ID if status == "已覆盖" else PENDING_HEADERS
    headers = _sheet_headers(ws, 2, len(expected))
    if headers != expected:
        errors.append(f"“{ws.title}”第 2 行表头必须为：{' | '.join(expected)}。")
        return []
    entries: list[dict[str, Any]] = []
    for row in range(3, ws.max_row + 1):
        cells = [
            ws.cell(row=row, column=column) for column in range(1, len(expected) + 1)
        ]
        values = [cell.value for cell in cells]
        if all(value is None or str(value).strip() == "" for value in values):
            continue
        for column, cell in enumerate(cells, start=1):
            if cell.data_type == "f":
                errors.append(f"“{ws.title}”第 {row} 行第 {column} 列包含公式，必须为纯文本。")
        if status == "已覆盖":
            trace_ids, case_errors = (
                parse_trace_id_list(_cell_text(values[4])) if _cell_text(values[4]) else ([], [])
            )
            errors.extend(f"“{ws.title}”第 {row} 行：{message}" for message in case_errors)
            entries.append(
                {
                    "id": _cell_text(values[0]),
                    "source": _cell_text(values[1]),
                    "location": _cell_text(values[2]),
                    "rule": _cell_text(values[3]),
                    "status": "已覆盖",
                    "case_ids": trace_ids,
                }
            )
        else:
            entries.append(
                {
                    "id": _cell_text(values[0]),
                    "source": _cell_text(values[1]),
                    "location": _cell_text(values[2]),
                    "status": "待确认",
                    "issue": _cell_text(values[3]),
                    "assumption": _cell_text(values[4]),
                    "question": _cell_text(values[5]),
                }
            )
    return entries


def _merge_duplicate_entries(
    entries: list[dict[str, Any]], warnings: list[str]
) -> list[dict[str, Any]]:
    """Drop rows repeating the same trace ID; conflicting IDs stay for the validator."""
    by_id: dict[str, dict[str, Any]] = {}
    result: list[dict[str, Any]] = []
    dropped = 0
    for entry in entries:
        trace_id = _text(entry.get("id"))
        if trace_id and trace_id in by_id:
            target = by_id[trace_id]
            if entry["status"] != "待确认":
                for case_id in entry.get("case_ids", []):
                    if case_id not in target.setdefault("case_ids", []):
                        target["case_ids"].append(case_id)
            dropped += 1
            continue
        if trace_id:
            by_id[trace_id] = entry
        result.append(entry)
    if dropped:
        warnings.append(f"合并了 {dropped} 条重复的追溯条目。")
    return result


def parse_excel_to_ir(
    excel_path: str | Path,
    sheet_name: str = "测试用例",
    merge_duplicates: bool = False,
) -> tuple[dict[str, Any], list[str], list[str]]:
    path = Path(excel_path)
    openpyxl = _require_openpyxl()
    if openpyxl is None:
        return {}, ["解析 xlsx 需要 openpyxl，请先安装: pip install openpyxl"], []
    try:
        workbook = openpyxl.load_workbook(path)
    except Exception as exc:  # corrupted zip, bad XML and permission errors alike
        return {}, [f"无法读取 Excel 文件: {path}（{exc}）"], []

    errors: list[str] = []
    warnings: list[str] = []
    if sheet_name not in workbook.sheetnames:
        return {}, [f"工作簿缺少主工作表“{sheet_name}”。"], []
    ws = workbook[sheet_name]
    headers = _sheet_headers(ws, 2, len(V2_HEADERS))

    meta = {
        "title": _cell_text(ws["A1"].value),
        "sheet_name": ws.title,
        **_parse_info_sheet(workbook, errors),
    }
    meta.setdefault("title", "")

    if headers == V2_HEADERS:
        cases = _parse_v2_case_sheet(ws, errors, warnings)
    else:
        return {}, [
            f"无法识别的工作表表头：{' | '.join(headers)}。"
        ], []

    entries: list[dict[str, Any]] = []
    if "需求追溯" in workbook.sheetnames:
        entries.extend(_parse_traceability_sheet(workbook["需求追溯"], "已覆盖", errors))
    else:
        errors.append("工作簿缺少“需求追溯”工作表。")
    if "待确认项" in workbook.sheetnames:
        entries.extend(_parse_traceability_sheet(workbook["待确认项"], "待确认", errors))
    if merge_duplicates:
        entries = _merge_duplicate_entries(entries, warnings)

    ir = {
        "schema_version": SCHEMA_VERSION,
        "meta": normalize_meta(meta),
        "traceability": entries,
        "cases": cases,
    }
    return ir, errors, warnings


# ---------------------------------------------------------------------------
# Merging artifacts
# ---------------------------------------------------------------------------


def _fill_missing_lists(target: dict[str, Any], source: dict[str, Any]) -> None:
    for field in ("pre_operations", "test_data", "steps", "expected_results", "post_operations"):
        if not target[field] and source.get(field):
            target[field] = list(source[field])


def merge_artifacts_to_ir(
    md_path: str | Path | None = None,
    excel_path: str | Path | None = None,
    primary: str = "markdown",
    sheet_name: str = "测试用例",
    allow_unassigned_ids: bool = False,
    strict_count: bool = True,
) -> tuple[dict[str, Any], list[str], list[str]]:
    errors: list[str] = []
    warnings: list[str] = []

    if not md_path and not excel_path:
        return {}, ["必须提供 Markdown 或 Excel 至少一个输入文件。"], []
    if primary not in ("markdown", "excel"):
        return {}, ["primary 只能为 markdown 或 excel。"], []

    md_ir: dict[str, Any] | None = None
    excel_ir: dict[str, Any] | None = None
    if md_path:
        md_ir, md_errors, md_warnings = parse_markdown_to_ir(
            md_path,
            allow_unassigned_ids=allow_unassigned_ids,
            strict_count=strict_count,
        )
        errors.extend(md_errors)
        warnings.extend(md_warnings)
    if excel_path:
        excel_ir, excel_errors, excel_warnings = parse_excel_to_ir(
            excel_path,
            sheet_name=sheet_name,
            merge_duplicates=True,
        )
        errors.extend(excel_errors)
        warnings.extend(excel_warnings)

    if not errors and md_ir is not None and not md_ir:
        errors.append(f"Markdown 解析失败: {md_path}")
    if not errors and excel_ir is not None and not excel_ir:
        errors.append(f"Excel 解析失败: {excel_path}")
    if errors:
        return {}, errors, warnings

    if md_ir is not None and excel_ir is not None:
        base = md_ir if primary == "markdown" else excel_ir
        other = excel_ir if primary == "markdown" else md_ir
    elif md_ir is not None:
        base, other = md_ir, None
    else:
        base, other = excel_ir, None
    if other is None:
        return base, errors, warnings

    other_cases = {case["id"]: case for case in other["cases"] if case["id"]}
    for case in base["cases"]:
        counterpart = other_cases.get(case["id"])
        if counterpart is None:
            continue
        if primary == "markdown":
            if not case["design_method"]:
                case["design_method"] = counterpart["design_method"]
        else:
            for field in ("module", "group", "polarity"):
                if not case[field] and counterpart.get(field):
                    case[field] = counterpart[field]
            # Markdown is authoritative when it is primary: a field deleted there
            # must stay deleted instead of being resurrected from the old Excel.
            _fill_missing_lists(case, counterpart)

    # 单源解析时的通用提示在合并后可能已不成立，按合并结果重算。
    stale = (
        "Markdown 不包含设计方法，导出 IR 后需从 Excel 合并或手工补充 design_method。",
        "Excel 不包含正向/反向信息，请在导出的 IR 中补充 polarity。",
    )
    warnings[:] = [warning for warning in warnings if warning not in stale]
    missing_design = [
        case["id"] or "无 ID 用例" for case in base["cases"] if not case["design_method"]
    ]
    if missing_design:
        warnings.append(
            "以下用例缺少 design_method，需补充后才能生成 Excel: "
            + "、".join(missing_design)
        )
    if primary == "excel":
        missing_polarity = [
            case["id"] or "无 ID 用例" for case in base["cases"] if not case["polarity"]
        ]
        if missing_polarity:
            warnings.append(
                "以下用例缺少 polarity（Excel 不承载该字段），需补充后才能渲染 Markdown: "
                + "、".join(missing_polarity)
            )

    if primary == "markdown" and excel_ir is not None:
        for field in ("coverage", "notes", "sheet_name"):
            if excel_ir["meta"].get(field):
                base["meta"][field] = excel_ir["meta"][field]
    elif primary == "excel" and md_ir is not None:
        for field in ("title", "requirement_name", "release_date"):
            if md_ir["meta"].get(field):
                base["meta"][field] = md_ir["meta"][field]

    return base, errors, warnings


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def _write_ir(path: Path, ir: dict[str, Any], force: bool) -> None:
    if path.exists() and not force:
        raise SystemExit(f"输出文件已存在: {path}。确认覆盖后请追加 --force。")
    path.parent.mkdir(parents=True, exist_ok=True)
    write_text_atomic(path, json.dumps(ir, ensure_ascii=False, indent=2) + "\n")


def _format_id_list(case_ids: list[str]) -> str:
    if len(case_ids) <= 30:
        return "、".join(case_ids) if case_ids else "（无）"
    head = "、".join(case_ids[:5])
    tail = "、".join(case_ids[-5:])
    return f"{head} … {tail}（共 {len(case_ids)} 条）"


def _inspect_summary(ir: dict[str, Any], source: str, ids: list[str]) -> list[str]:
    meta = ir.get("meta") or {}
    cases = [case for case in ir.get("cases") or [] if isinstance(case, dict)]
    entries = [entry for entry in ir.get("traceability") or [] if isinstance(entry, dict)]
    covered = [entry for entry in entries if entry.get("status") == "已覆盖"]
    pending = [entry for entry in entries if entry.get("status") == "待确认"]

    lines = [f"=== {source} ==="]
    lines.append(f"需求名称：{meta.get('requirement_name') or '（空）'}")
    lines.append(f"上线日期：{meta.get('release_date') or '（空）'}")
    lines.append(f"覆盖范围：{meta.get('coverage') or '（空）'}")
    lines.append(f"用例总数：{len(cases)}")
    lines.append(f"用例 ID：{_format_id_list([case.get('id', '') for case in cases])}")

    modules = Counter(case.get("module", "") for case in cases)
    types = Counter(case.get("type", "") for case in cases)
    lines.append(
        "按模块："
        + ("、".join(f"{name or '（未标注）'}({count})" for name, count in modules.items()) or "（无）")
    )
    lines.append(
        "按类型："
        + ("、".join(f"{name or '（未标注）'}({count})" for name, count in types.items()) or "（无）")
    )
    lines.append(f"已覆盖追溯：{len(covered)} 条；待确认：{len(pending)} 条")
    if pending:
        for entry in pending:
            lines.append(f"待确认 {entry.get('id', '')}：{entry.get('question', '')}")

    if ids:
        by_id = {case.get("id", ""): case for case in cases}
        for wanted in ids:
            case = by_id.get(wanted)
            if case is None:
                lines.append(f"未找到用例: {wanted}")
                continue
            lines.append(f"--- {wanted} ---")
            lines.append(f"标题：{case.get('title', '')}")
            lines.append(f"类型：{case.get('type', '')}；设计方法：{case.get('design_method') or '（空）'}")
            lines.append(f"优先级：{case.get('priority', '')}")
            raw_refs = case.get("traceability_refs")
            refs = (
                [str(ref) for ref in raw_refs if ref is not None]
                if isinstance(raw_refs, (list, tuple))
                else []
            )
            lines.append("追溯引用：" + ("、".join(refs) or "（无）"))
            for label, key in (
                ("测试步骤", "steps"),
                ("预期结果", "expected_results"),
                ("后置操作", "post_operations"),
            ):
                items = case.get(key) or []
                lines.append(f"{label}：" + ("；".join(items) if items else "（空）"))
    return lines


def inspect_artifacts(
    markdown_path: str | Path | None = None,
    excel_path: str | Path | None = None,
    baseline_path: str | Path | None = None,
    sheet_name: str = "测试用例",
    ids: list[str] | None = None,
) -> tuple[list[str], list[str]]:
    """Build a read-only summary of Markdown / Excel / IR artifacts."""
    if markdown_path is None and excel_path is None and baseline_path is None:
        return [], ["必须提供 --markdown、--excel 或 --baseline 至少一个。"]

    lines: list[str] = []
    errors: list[str] = []
    wanted = [value for value in (ids or []) if value]

    if baseline_path is not None:
        try:
            ir = load_baseline(baseline_path)
        except ValueError as exc:
            errors.append(str(exc))
        else:
            lines.extend(_inspect_summary(ir, "Baseline IR", wanted))
            lines.append("")

    if markdown_path is not None:
        ir, md_errors, _ = parse_markdown_to_ir(markdown_path, strict_count=False)
        if md_errors:
            errors.extend(f"Markdown: {message}" for message in md_errors)
        else:
            lines.extend(_inspect_summary(ir, "Markdown", wanted))
            lines.append("")

    if excel_path is not None:
        ir, excel_errors, _ = parse_excel_to_ir(excel_path, sheet_name=sheet_name)
        if excel_errors:
            errors.extend(f"Excel: {message}" for message in excel_errors)
        else:
            import openpyxl

            workbook = openpyxl.load_workbook(excel_path, read_only=True)
            try:
                sheet_lines = _inspect_summary(ir, "Excel", wanted)
                sheet_lines.insert(1, "工作表：" + "、".join(workbook.sheetnames))
            finally:
                workbook.close()
            lines.extend(sheet_lines)
            lines.append("")

    while lines and not lines[-1]:
        lines.pop()
    return lines, errors


def main(argv: list[str] | None = None) -> int:
    configure_utf8_stdio()
    parser = argparse.ArgumentParser(description="校验、导出与维护合并 IR。")
    subparsers = parser.add_subparsers(dest="command", required=True)

    check = subparsers.add_parser("check", help="校验 IR 文件")
    check.add_argument("--baseline", required=True)
    check.add_argument("--require-excel-fields", action="store_true")

    export = subparsers.add_parser("export", help="从现有交付物导出 IR")
    export.add_argument("--markdown")
    export.add_argument("--excel")
    export.add_argument("--sheet-name", default="测试用例")
    export.add_argument("--output", required=True)
    export.add_argument("--force", action="store_true")

    assign = subparsers.add_parser("assign-ids", help="为 IR 中缺失的用例补号")
    assign.add_argument("--baseline", required=True)
    assign.add_argument("--id-prefix")

    renumber = subparsers.add_parser("renumber", help="定稿前重排全部用例 ID")
    renumber.add_argument("--baseline", required=True)
    renumber.add_argument("--id-prefix")

    inspect = subparsers.add_parser("inspect", help="只读打印产物摘要")
    inspect.add_argument("--markdown")
    inspect.add_argument("--excel")
    inspect.add_argument("--baseline")
    inspect.add_argument("--sheet-name", default="测试用例")
    inspect.add_argument("--ids", help="逗号分隔的用例 ID，打印明细")

    args = parser.parse_args(argv)

    try:
        if args.command == "inspect":
            ids = [value.strip() for value in (args.ids or "").split(",")]
            lines, errors = inspect_artifacts(
                markdown_path=args.markdown,
                excel_path=args.excel,
                baseline_path=args.baseline,
                sheet_name=args.sheet_name,
                ids=[value for value in ids if value],
            )
            for line in lines:
                print(line)
            if errors:
                for error in errors:
                    print(error, file=sys.stderr)
                return 1
            return 0

        if args.command == "check":
            errors = validate_baseline(
                _read_json(args.baseline),
                require_design_method=args.require_excel_fields,
            )
            if errors:
                print("\n".join(errors), file=sys.stderr)
                return 1
            ir = load_baseline(args.baseline)
            covered = sum(1 for entry in ir["traceability"] if entry.get("status") == "已覆盖")
            pending = sum(1 for entry in ir["traceability"] if entry.get("status") == "待确认")
            print(
                f"校验通过：共 {len(ir['cases'])} 条用例，"
                f"已覆盖追溯 {covered} 条，待确认 {pending} 条。"
            )
            return 0

        if args.command == "export":
            ir, errors, warnings = merge_artifacts_to_ir(
                md_path=args.markdown,
                excel_path=args.excel,
                sheet_name=args.sheet_name,
            )
            if errors:
                print("\n".join(errors), file=sys.stderr)
                return 1
            _write_ir(Path(args.output).expanduser().resolve(), ir, args.force)
            for warning in warnings:
                print(f"提示：{warning}", file=sys.stderr)
            print(f"已导出 IR：{args.output}")
            return 0

        if args.command == "assign-ids":
            ir = load_baseline(args.baseline)
            assigned = assign_missing_ids(ir["cases"], args.id_prefix)
            _write_ir(Path(args.baseline).expanduser().resolve(), ir, force=True)
            if assigned:
                print(f"已补号 {len(assigned)} 条：{', '.join(assigned)}")
            else:
                print("没有需要补号的用例。")
            return 0

        if args.command == "renumber":
            ir = load_baseline(args.baseline)
            mapping = renumber_cases(ir["cases"], ir["traceability"], args.id_prefix)
            _write_ir(Path(args.baseline).expanduser().resolve(), ir, force=True)
            print(f"已重排 {len(mapping)} 条用例 ID。")
            return 0
    except (ValueError, FileLockedError) as exc:
        print(str(exc), file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
