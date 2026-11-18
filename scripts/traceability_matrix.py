"""Load and validate the structured requirement traceability matrix."""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

STATUSES = {"待覆盖", "已覆盖", "待确认"}
PENDING_FIELDS = ("issue", "assumption", "question")
REQUIRED_FIELDS = ("id", "source", "location", "rule", "status")
TRACEABILITY_PATTERN = re.compile(
    r"^\[需求追溯\]\s+来源：(.+?)｜定位：(.+?)｜规则：(.+?)$"
)


def traceability_text(entry: dict[str, Any]) -> str:
    return (
        f"[需求追溯] 来源：{entry['source']}｜定位：{entry['location']}｜规则：{entry['rule']}"
    )


def entries_for_case_id(
    matrix: dict[str, dict[str, Any]],
    case_id: str,
) -> list[dict[str, Any]]:
    return [
        entry
        for entry in matrix.values()
        if case_id in entry["case_ids"]
    ]


def entries_for_case_title(
    matrix: dict[str, dict[str, Any]],
    case_title: str,
) -> list[dict[str, Any]]:
    return [
        entry
        for entry in matrix.values()
        if case_title in entry["case_titles"]
    ]


def final_matrix_errors(matrix: dict[str, dict[str, Any]]) -> list[str]:
    pending_coverage = [
        entry["id"]
        for entry in matrix.values()
        if entry["status"] == "待覆盖"
    ]
    if not pending_coverage:
        return []
    return [
        "需求追溯矩阵存在待覆盖条目，暂不能核对用例关联: "
        + "、".join(pending_coverage)
        + "。请完成各格式关联后再执行最终校验。"
    ]


def load_traceability_matrix(path: str | Path) -> tuple[dict[str, dict[str, Any]], list[str]]:
    matrix_path = Path(path)
    if not matrix_path.is_file():
        return {}, [f"需求追溯矩阵不存在: {matrix_path}"]

    try:
        payload = json.loads(matrix_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        return {}, [f"无法读取需求追溯矩阵: {exc}"]

    entries = payload.get("entries") if isinstance(payload, dict) else None
    if not isinstance(entries, list) or not entries:
        return {}, ["需求追溯矩阵必须包含非空 entries 数组。"]

    result: dict[str, dict[str, Any]] = {}
    errors: list[str] = []
    traceability_references: dict[str, str] = {}
    for index, entry in enumerate(entries, start=1):
        if not isinstance(entry, dict):
            errors.append(f"需求追溯矩阵第 {index} 项必须是对象。")
            continue
        missing = [field for field in REQUIRED_FIELDS if field not in entry]
        if missing:
            errors.append(f"需求追溯矩阵第 {index} 项缺少字段: {', '.join(missing)}。")
            continue

        trace_id = str(entry["id"]).strip()
        text_fields = ("source", "location", "rule")
        if not trace_id or any(not str(entry[field]).strip() for field in text_fields):
            errors.append(f"需求追溯矩阵第 {index} 项的 ID、来源、定位和规则不能为空。")
            continue
        if trace_id in result:
            errors.append(f"需求追溯矩阵中的追溯 ID 重复: '{trace_id}'。")
            continue

        normalized = {
            "id": trace_id,
            "source": str(entry["source"]).strip(),
            "location": str(entry["location"]).strip(),
            "rule": str(entry["rule"]).strip(),
            "status": str(entry["status"]).strip(),
            "case_ids": [str(value).strip() for value in entry["case_ids"] if str(value).strip()]
            if isinstance(entry.get("case_ids"), list)
            else [],
            "case_titles": [str(value).strip() for value in entry["case_titles"] if str(value).strip()]
            if isinstance(entry.get("case_titles"), list)
            else [],
        }
        status = normalized["status"]
        if status not in STATUSES:
            errors.append(
                f"需求追溯矩阵第 {index} 项的 status 必须为“待覆盖”“已覆盖”或“待确认”。"
            )
            continue
        if len(set(normalized["case_ids"])) != len(normalized["case_ids"]):
            errors.append(
                f"需求追溯矩阵第 {index} 项的 case_ids 不得重复。"
            )
        if len(set(normalized["case_titles"])) != len(normalized["case_titles"]):
            errors.append(
                f"需求追溯矩阵第 {index} 项的 case_titles 不得重复。"
            )
        if status == "待覆盖":
            if normalized["case_ids"] or normalized["case_titles"]:
                errors.append(f"需求追溯矩阵第 {index} 项待覆盖时，不得关联用例。")
            if any(str(entry.get(field) or "").strip() for field in PENDING_FIELDS):
                errors.append(f"需求追溯矩阵第 {index} 项待覆盖时，不得填写待确认字段。")
        elif status == "已覆盖":
            if not normalized["case_ids"] and not normalized["case_titles"]:
                errors.append(
                    f"需求追溯矩阵第 {index} 项已覆盖时，case_ids 和 case_titles 至少填写一项。"
                )
            if any(str(entry.get(field) or "").strip() for field in PENDING_FIELDS):
                errors.append(f"需求追溯矩阵第 {index} 项已覆盖时，不得保留待确认字段。")
        else:
            missing_pending_fields = [
                field
                for field in PENDING_FIELDS
                if not str(entry.get(field) or "").strip()
            ]
            if missing_pending_fields:
                errors.append(
                    f"需求追溯矩阵第 {index} 项待确认时，"
                    f"issue、assumption、question 均不能为空，缺少: {', '.join(missing_pending_fields)}。"
                )
            if normalized["case_ids"] or normalized["case_titles"]:
                errors.append(f"需求追溯矩阵第 {index} 项待确认时，不得关联用例。")
        traceability = traceability_text(normalized)
        other_id = traceability_references.get(traceability)
        if other_id:
            errors.append(
                f"需求追溯矩阵中的追溯文本同时属于 '{other_id}' 和 '{trace_id}'，"
                "请使用不同的来源、定位或规则。"
            )
        else:
            traceability_references[traceability] = trace_id
        result[trace_id] = normalized

    return result, errors
