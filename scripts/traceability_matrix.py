"""Validate the structured requirement traceability matrix."""

from __future__ import annotations

import re
from typing import Any

STATUSES = {"待覆盖", "已覆盖", "待确认"}
PENDING_FIELDS = ("issue", "assumption", "question")
TRACE_ID_INVALID_PATTERN = re.compile(r"[\s,\[\]，]")


def _text_field(value: Any) -> str:
    return str(value).strip() if value is not None else ""


def _has_newline(value: str) -> bool:
    return "\r" in value or "\n" in value


def parse_trace_id_list(text: Any) -> tuple[list[str], list[str]]:
    """Parse comma-separated trace IDs, accepting both ASCII and full-width commas."""
    raw = _text_field(text)
    if not raw:
        return [], ["追溯 ID 不能为空。"]

    ids: list[str] = []
    errors: list[str] = []
    for part in re.split(r"[,，]", raw):
        trace_id = part.strip()
        if not trace_id:
            errors.append("追溯 ID 列表中存在空项。")
            continue
        if TRACE_ID_INVALID_PATTERN.search(trace_id):
            errors.append(f"追溯 ID '{trace_id}' 含非法字符（空白、逗号或方括号）。")
            continue
        if trace_id in ids:
            errors.append(f"追溯 ID 重复: '{trace_id}'。")
            continue
        ids.append(trace_id)
    return ids, errors


def format_trace_id_list(entries_or_ids: Any) -> str:
    """Join trace entries (dicts with ``id``) or plain IDs with ``, ``."""
    items: list[str] = []
    for item in entries_or_ids or []:
        value = item.get("id") if isinstance(item, dict) else item
        text = _text_field(value)
        if text:
            items.append(text)
    return ", ".join(items)


def _normalize_case_ids(value: Any) -> tuple[list[str], list[str]]:
    if value is None:
        return [], []
    if not isinstance(value, (list, tuple)):
        return [], ["case_ids 必须为数组。"]
    case_ids: list[str] = []
    errors: list[str] = []
    for item in value:
        text = _text_field(item)
        if not text:
            errors.append("case_ids 中存在空值。")
            continue
        if text in case_ids:
            errors.append(f"case_ids 重复: '{text}'。")
            continue
        case_ids.append(text)
    return case_ids, errors


def validate_traceability_entries(
    raw_entries: Any,
    allow_unassigned_ids: bool = False,
    allow_unsettled_links: bool = False,
) -> tuple[dict[str, dict[str, Any]], list[str]]:
    """Validate traceability entries from either an IR array or parsed tables.

    ``allow_unsettled_links`` is the draft/settlement mode: links on 待覆盖 and
    已覆盖 entries are not judged here because ``sync_traceability_from_cases``
    recomputes them from each case's ``traceability_refs``.
    """
    if isinstance(raw_entries, dict):
        candidates = list(raw_entries.values())
    elif isinstance(raw_entries, (list, tuple)):
        candidates = list(raw_entries)
    else:
        return {}, ["追溯条目必须为数组。"]

    if not candidates:
        return {}, ["追溯条目不能为空。"]

    result: dict[str, dict[str, Any]] = {}
    errors: list[str] = []
    covered_keys: dict[tuple[str, str, str], str] = {}
    pending_keys: dict[tuple[str, str, str], str] = {}

    for index, entry in enumerate(candidates, start=1):
        label = f"追溯条目第 {index} 项"
        if not isinstance(entry, dict):
            errors.append(f"{label}必须是对象。")
            continue
        if "case_titles" in entry:
            errors.append(
                f"{label}包含不支持字段 case_titles，"
                "请改用 case_ids 或用例正文的引用行。"
            )
            continue

        trace_id = _text_field(entry.get("id"))
        if not trace_id:
            errors.append(f"{label}的 id 不能为空。")
            continue
        if TRACE_ID_INVALID_PATTERN.search(trace_id):
            errors.append(
                f"{label}的 id '{trace_id}' 含非法字符（空白、逗号或方括号）。"
            )
            continue
        if trace_id in result:
            errors.append(f"追溯 ID 重复: '{trace_id}'。")
            continue

        source = _text_field(entry.get("source"))
        location = _text_field(entry.get("location"))
        if not source or not location:
            errors.append(f"{label}（{trace_id}）的 source 与 location 不能为空。")
            continue
        if _has_newline(source) or _has_newline(location):
            errors.append(f"{label}（{trace_id}）的 source 与 location 不得包含换行。")
            continue

        status = _text_field(entry.get("status"))
        if status not in STATUSES:
            errors.append(
                f"{label}（{trace_id}）的 status 必须为“待覆盖”“已覆盖”或“待确认”。"
            )
            continue

        case_ids, case_errors = _normalize_case_ids(entry.get("case_ids"))
        if allow_unsettled_links and status in ("待覆盖", "已覆盖"):
            case_errors = [message for message in case_errors if "必须为数组" in message]
        if case_errors:
            errors.extend(
                f"{label}（{trace_id}）{message}" for message in case_errors
            )
            if status == "待确认" and case_ids:
                errors.append(f"{label}（{trace_id}）为待确认条目，不得关联用例。")
            continue

        normalized: dict[str, Any] = {
            "id": trace_id,
            "source": source,
            "location": location,
            "status": status,
            "case_ids": [],
        }

        if status == "待确认":
            pending_values = {
                field: _text_field(entry.get(field)) for field in PENDING_FIELDS
            }
            missing = [field for field, value in pending_values.items() if not value]
            if missing:
                errors.append(
                    f"{label}（{trace_id}）为待确认条目，"
                    f"issue、assumption、question 均不能为空，缺少: {', '.join(missing)}。"
                )
                continue
            newline_fields = [
                field for field, value in pending_values.items() if _has_newline(value)
            ]
            if newline_fields:
                errors.append(
                    f"{label}（{trace_id}）的 {', '.join(newline_fields)} 不得包含换行。"
                )
                continue
            if case_ids:
                errors.append(f"{label}（{trace_id}）为待确认条目，不得关联用例。")
                continue
            key = (source, location, pending_values["issue"])
            other_id = pending_keys.get(key)
            if other_id:
                errors.append(
                    f"{label}（{trace_id}）与 '{other_id}' 的来源、定位和未明确项重复，"
                    "请合并为同一待确认条目。"
                )
                continue
            pending_keys[key] = trace_id
            normalized.update(pending_values)
        else:
            rule = _text_field(entry.get("rule"))
            if not rule:
                errors.append(f"{label}（{trace_id}）的 rule 不能为空。")
                continue
            if _has_newline(rule):
                errors.append(f"{label}（{trace_id}）的 rule 不得包含换行。")
                continue
            filled_pending = [
                field for field in PENDING_FIELDS if _text_field(entry.get(field))
            ]
            if filled_pending:
                errors.append(
                    f"{label}（{trace_id}）为{status}条目，"
                    f"不得填写待确认字段: {', '.join(filled_pending)}。"
                )
                continue
            key = (source, location, rule)
            other_id = covered_keys.get(key)
            if other_id:
                errors.append(
                    f"{label}（{trace_id}）与 '{other_id}' 的来源、定位和规则重复，"
                    "请合并为同一追溯条目。"
                )
                continue
            covered_keys[key] = trace_id
            if (
                status == "已覆盖"
                and not case_ids
                and not allow_unassigned_ids
                and not allow_unsettled_links
            ):
                errors.append(
                    f"{label}（{trace_id}）为已覆盖条目，必须至少关联一个用例。"
                )
                continue
            if status == "待覆盖" and case_ids and not allow_unsettled_links:
                errors.append(f"{label}（{trace_id}）为待覆盖条目，不得关联用例。")
                continue
            normalized["rule"] = rule
            normalized["case_ids"] = case_ids

        result[trace_id] = normalized

    return result, errors


def describe_traceability_mismatch(
    expected_ids: Any,
    actual_ids: Any,
) -> str:
    """Describe the difference between expected and actual trace ID sequences."""
    expected = [_text_field(item) for item in expected_ids or []]
    actual = [_text_field(item) for item in actual_ids or []]
    expected = [item for item in expected if item]
    actual = [item for item in actual if item]
    if expected == actual:
        return ""

    missing = [item for item in expected if item not in actual]
    extra = [item for item in actual if item not in expected]
    if not missing and not extra:
        if sorted(expected) == sorted(actual):
            return (
                f"追溯关联顺序不一致：期望=[{', '.join(expected)}]，"
                f"实际=[{', '.join(actual)}]。"
            )
        return (
            f"追溯关联不一致：期望=[{', '.join(expected)}]，"
            f"实际=[{', '.join(actual)}]；"
            f"期望 {len(expected)} 个、实际 {len(actual)} 个，存在重复项。"
        )

    details: list[str] = []
    if missing:
        details.append(f"缺失: {', '.join(missing)}")
    if extra:
        details.append(f"多余: {', '.join(extra)}")
    return (
        f"追溯关联不一致：期望=[{', '.join(expected)}]，"
        f"实际=[{', '.join(actual)}]；{'；'.join(details)}。"
    )


def sync_traceability_from_cases(
    traceability: Any,
    cases: Any,
) -> tuple[dict[str, dict[str, Any]], list[str], list[str]]:
    """Derive ``case_ids`` from each case's ``traceability_refs`` and settle statuses."""
    entries: dict[str, dict[str, Any]] = {}
    errors: list[str] = []
    warnings: list[str] = []
    if isinstance(traceability, dict):
        source_entries: list[Any] = []
        for key, value in traceability.items():
            if isinstance(value, dict) and not _text_field(value.get("id")):
                value = {**value, "id": key}
            source_entries.append(value)
    elif isinstance(traceability, (list, tuple)):
        source_entries = list(traceability)
    else:
        return {}, ["追溯条目必须为数组。"], []

    for index, item in enumerate(source_entries, start=1):
        if not isinstance(item, dict):
            errors.append(f"第 {index} 个追溯条目必须是对象。")
            continue
        entry = dict(item)
        trace_id = _text_field(entry.get("id"))
        entry["id"] = trace_id
        entry["status"] = _text_field(entry.get("status"))
        entry["case_ids"] = []
        if not trace_id:
            errors.append(f"第 {index} 个追溯条目缺少 id，无法派生用例关联。")
            continue
        if trace_id in entries:
            errors.append(f"追溯 ID 重复: '{trace_id}'。")
            continue
        entries[trace_id] = entry

    seen_case_ids: set[str] = set()

    for index, case in enumerate(cases or [], start=1):
        if not isinstance(case, dict):
            errors.append(f"第 {index} 条用例必须是对象。")
            continue
        case_id = _text_field(case.get("id"))
        if not case_id:
            errors.append(f"第 {index} 条用例缺少用例 ID，无法派生追溯关联。")
            continue
        if case_id in seen_case_ids:
            errors.append(f"用例 ID 重复: '{case_id}'。")
            continue
        seen_case_ids.add(case_id)

        refs = case.get("traceability_refs")
        if refs is None:
            refs = []
        if not isinstance(refs, (list, tuple)):
            errors.append(f"用例 '{case_id}' 的 traceability_refs 必须为数组。")
            continue
        for ref in refs:
            trace_id = _text_field(ref)
            if not trace_id:
                errors.append(f"用例 '{case_id}' 的 traceability_refs 中存在空值。")
                continue
            entry = entries.get(trace_id)
            if entry is None:
                errors.append(
                    f"用例 '{case_id}' 引用了不存在的追溯 ID '{trace_id}'。"
                )
                continue
            if entry.get("status") == "待确认":
                errors.append(
                    f"用例 '{case_id}' 不得引用待确认条目 '{trace_id}'。"
                )
                continue
            if case_id not in entry["case_ids"]:
                entry["case_ids"].append(case_id)

    for entry in entries.values():
        if entry.get("status") == "待覆盖" and entry["case_ids"]:
            entry["status"] = "已覆盖"
        elif entry.get("status") == "已覆盖" and not entry["case_ids"]:
            entry["status"] = "待覆盖"
            warnings.append(
                f"追溯条目 '{entry['id']}' 失去全部关联用例，已回退为待覆盖，"
                "请补充用例或移除该条目。"
            )

    return entries, errors, warnings
