"""Load and validate the structured requirement traceability matrix."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

STATUSES = {"已覆盖", "待确认"}
PENDING_FIELDS = ("issue", "assumption", "question")
REQUIRED_FIELDS = ("id", "source", "location", "rule", "status")


def traceability_text(entry: dict[str, Any]) -> str:
    return (
        f"[需求追溯] 来源：{entry['source']}｜定位：{entry['location']}｜规则：{entry['rule']}"
    )


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
    case_id_references: dict[str, str] = {}
    case_title_references: dict[str, str] = {}
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
                f"需求追溯矩阵第 {index} 项的 status 必须为“已覆盖”或“待确认”。"
            )
            continue
        if status == "已覆盖" and not normalized["case_ids"] and not normalized["case_titles"]:
            errors.append(
                f"需求追溯矩阵第 {index} 项已覆盖时，case_ids 和 case_titles 至少填写一项。"
            )
        elif status == "待确认":
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
        if status == "已覆盖" and any(
            str(entry.get(field) or "").strip() for field in PENDING_FIELDS
        ):
            errors.append(f"需求追溯矩阵第 {index} 项已覆盖时，不得保留待确认字段。")
        for field, references, label in (
            ("case_ids", case_id_references, "用例 ID"),
            ("case_titles", case_title_references, "Markdown 用例标题"),
        ):
            for reference in normalized[field]:
                other_id = references.get(reference)
                if other_id:
                    errors.append(
                        f"需求追溯矩阵中的{label} '{reference}' 同时属于 '{other_id}' 和 '{trace_id}'。"
                    )
                else:
                    references[reference] = trace_id
        result[trace_id] = normalized

    return result, errors
