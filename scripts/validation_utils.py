"""Shared validation helpers used by the generator and validators."""

from __future__ import annotations

import re
from datetime import date

DATE_PLACEHOLDER = "未指定日期"
DATE_PATTERN = re.compile(r"\d{4}-\d{2}-\d{2}")


def is_valid_date_or_placeholder(value: object) -> bool:
    """Return True for a valid YYYY-MM-DD date or the documented placeholder."""
    text = str(value).strip() if value is not None else ""
    if text == DATE_PLACEHOLDER:
        return True
    if not DATE_PATTERN.fullmatch(text):
        return False
    try:
        date.fromisoformat(text)
    except ValueError:
        return False
    return True
