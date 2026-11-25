"""Shared validation helpers used by the generator and validators."""

from __future__ import annotations

import math
import os
import re
import shutil
import stat
import sys
import tempfile
import unicodedata
from datetime import date
from pathlib import Path
from typing import Any

DATE_PLACEHOLDER = "未指定日期"
DATE_PATTERN = re.compile(r"\d{4}-\d{2}-\d{2}")

ALLOWED_CASE_TYPES = (
    "功能测试",
    "回归测试",
    "性能测试",
    "安全测试",
    "兼容性测试",
    "易用性测试",
    "可靠性测试",
)
ALLOWED_DESIGN_METHODS = (
    "等价类划分",
    "边界值分析",
    "判定表法",
    "状态迁移法",
    "错误推测法",
    "正交试验",
    "场景分析法",
)
ALLOWED_PRIORITIES = ("P0", "P1", "P2", "P3", "P4")
ALLOWED_POLARITIES = ("正向", "反向")

LIST_ITEM_PREFIX_PATTERN = re.compile(r"^\s*(?:\d+\.|[-*])\s+")
DEFAULT_COLUMN_WIDTH = 40
MIN_ROW_HEIGHT = 24
MAX_ROW_HEIGHT = 409
ROW_HEIGHT_PER_LINE = 16
ROW_HEIGHT_PADDING = 8


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


def configure_utf8_stdio() -> None:
    """Force UTF-8 on both stdout and stderr for Windows console and pipes."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except (OSError, ValueError):
            continue


class FileLockedError(OSError):
    """Raised when a target file cannot be written because another process holds it."""


def is_file_locked_error(exc: OSError) -> bool:
    """Return True for the errors Windows raises when a file is opened elsewhere.

    ``winerror`` 32/33 are sharing violations; ``winerror`` 5 is access denied
    (read-only attribute or ACL) and is reported as "not writable" instead.
    """
    winerror = getattr(exc, "winerror", None)
    if winerror is not None:
        return winerror in (32, 33)
    return isinstance(exc, PermissionError)


def _file_locked_message(path: Path) -> str:
    return (
        f"文件被占用，无法写入：{path}。"
        "请先关闭正在打开该文件的程序（如 Excel、WPS 或编辑器）后重试。"
    )


def _file_unwritable_message(path: Path) -> str:
    return f"文件不可写（只读或权限不足）：{path}。请解除只读属性或检查写入权限后重试。"


def _file_denied_message(path: Path) -> str:
    return (
        f"无法写入：{path}（文件可能被其他程序占用或权限不足）。"
        "请先关闭占用程序（如 Excel、WPS）并检查写入权限后重试。"
    )


def locked_write_error(exc: OSError, path: Path) -> FileLockedError:
    """Map a failed write to the matching Chinese lock/denied message."""
    if getattr(exc, "winerror", None) == 5:
        return FileLockedError(_file_denied_message(path))
    return FileLockedError(_file_locked_message(path))


def readonly_write_error(exc: OSError, path: Path) -> PermissionError | None:
    """Return a read-only/permission error when the target itself is unwritable."""
    if isinstance(exc, PermissionError) and path.exists() and _is_readonly_or_unwritable(path):
        return PermissionError(_file_unwritable_message(path))
    return None


def _parent_unwritable_message(path: Path) -> str:
    return f"输出目录不可写（权限不足）：{path}。请检查目录权限后重试。"


def _is_readonly_or_unwritable(path: Path) -> bool:
    """Return True when the existing file cannot be opened for writing."""
    try:
        mode = path.stat().st_mode
    except OSError:
        return True
    if os.name == "nt":
        return not bool(mode & stat.S_IWRITE)
    return not os.access(path, os.W_OK)


def _is_parent_writable(parent: Path) -> bool:
    """Return True when ``parent`` exists and accepts new files."""
    if not parent.is_dir():
        return False
    if os.name != "nt":
        return os.access(parent, os.W_OK)
    try:
        fd, probe = tempfile.mkstemp(prefix=".write-probe.", dir=str(parent))
    except OSError:
        return False
    os.close(fd)
    try:
        os.unlink(probe)
    except OSError:
        pass
    return True


def _check_targets_writable(targets: list[Path]) -> list[str]:
    """Pre-flight write checks so a locked target never disturbs the others."""
    errors: list[str] = []
    for target in targets:
        if not target.exists():
            if not _is_parent_writable(target.parent):
                errors.append(_parent_unwritable_message(target.parent))
            continue
        if _is_readonly_or_unwritable(target):
            errors.append(_file_unwritable_message(target))
            continue
        if os.name != "nt":
            # POSIX has no portable lock probe; os.replace plus rollback covers it.
            continue
        try:
            with target.open("r+b"):
                pass
        except OSError as exc:
            # The file is already known to be writable, so a sharing violation or a
            # winerror-less PermissionError means another process holds it; winerror 5
            # is an ACL denial, and anything else (e.g. ENOSPC) keeps its own text.
            winerror = getattr(exc, "winerror", None)
            if winerror in (32, 33) or (winerror is None and isinstance(exc, PermissionError)):
                errors.append(_file_locked_message(target))
            elif winerror == 5:
                errors.append(_file_denied_message(target))
            else:
                errors.append(f"无法写入：{target}（{exc}）")
    return errors


def replace_file_atomic(temp_path: str | Path, target_path: str | Path) -> None:
    """Replace ``target_path`` with ``temp_path`` atomically, cleaning up on failure."""
    temp_path = Path(temp_path)
    target_path = Path(target_path)
    try:
        os.replace(temp_path, target_path)
    except OSError as exc:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass
        if isinstance(exc, PermissionError) and not target_path.exists():
            if not _is_parent_writable(target_path.parent):
                raise PermissionError(
                    f"输出目录不可写（权限不足），无法写入：{target_path}"
                ) from exc
            raise
        readonly_error = readonly_write_error(exc, target_path)
        if readonly_error is not None:
            raise readonly_error from exc
        if is_file_locked_error(exc) or getattr(exc, "winerror", None) == 5:
            raise locked_write_error(exc, target_path) from exc
        raise


def write_text_atomic(path: str | Path, text: str) -> None:
    """Write UTF-8 text through a sibling temp file and an atomic replace."""
    path = Path(path)
    handle = tempfile.NamedTemporaryFile(
        mode="w",
        encoding="utf-8",
        newline="\n",
        delete=False,
        dir=str(path.parent),
        prefix=f".{path.name}.",
        suffix=".tmp",
    )
    temp_path = Path(handle.name)
    try:
        with handle:
            handle.write(text)
        replace_file_atomic(temp_path, path)
    except BaseException:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise


def sibling_temp_path(path: str | Path) -> Path:
    """Create an empty temporary file next to ``path`` for staged output."""
    target = Path(path)
    fd, temp_name = tempfile.mkstemp(
        prefix=f".{target.name}.",
        suffix=target.suffix or ".tmp",
        dir=str(target.parent),
    )
    os.close(fd)
    return Path(temp_name)


def copy_to_temp(path: str | Path) -> Path:
    """Copy ``path`` to a sibling temporary file, removing it if the copy fails."""
    source = Path(path)
    fd, temp_name = tempfile.mkstemp(
        prefix=f".{source.name}.",
        suffix=source.suffix or ".tmp",
        dir=str(source.parent),
    )
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        shutil.copy2(source, temp_path)
    except BaseException:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise
    return temp_path


def commit_files_atomic(staged: list[tuple[Path, Path]]) -> list[str]:
    """Replace each target with its staged file, rolling back if a later write fails.

    ``staged`` pairs are ``(target, temp_file)`` in commit order. Each existing
    target is backed up first; a failed replacement restores everything already
    committed and returns Chinese error messages. Staged and backup files are
    always cleaned up.
    """
    backups: dict[Path, Path | None] = {}
    committed: list[Path] = []
    try:
        preflight_errors = _check_targets_writable([target for target, _temp in staged])
        if preflight_errors:
            return preflight_errors
        for target, _temp in staged:
            try:
                backups[target] = copy_to_temp(target) if target.exists() else None
            except OSError as exc:
                return [f"备份文件失败: {target}（{exc}）"]
        for target, temp in staged:
            try:
                replace_file_atomic(temp, target)
            except OSError as exc:
                if isinstance(exc, FileLockedError):
                    detail = str(exc)
                elif (
                    is_file_locked_error(exc)
                    and exc.errno is not None
                    and target.exists()
                    and not _is_readonly_or_unwritable(target)
                ):
                    detail = _file_locked_message(target)
                else:
                    detail = str(exc)
                messages = [f"写入文件失败: {target}（{detail}）"]
                for done in reversed(committed):
                    backup = backups[done]
                    try:
                        if backup is None:
                            done.unlink(missing_ok=True)
                        else:
                            os.replace(backup, done)
                            backups[done] = None
                    except OSError as rollback_exc:
                        messages.append(
                            f"回滚 {done} 失败: {rollback_exc}"
                            + (f"；备份保留在 {backup}，请手工恢复。" if backup else "。")
                        )
                return messages
            committed.append(target)
        return []
    finally:
        for _target, temp in staged:
            try:
                Path(temp).unlink(missing_ok=True)
            except OSError:
                pass
        for backup in backups.values():
            if backup is None:
                continue
            try:
                backup.unlink(missing_ok=True)
            except OSError:
                pass


def normalize_list_items(value: Any) -> list[str]:
    """Split a field into non-empty lines, dropping redundant list markers.

    Markers are removed only when text remains after the marker, so values such
    as ``-1+1`` or ``1.5 秒`` keep their leading characters.
    """
    if value is None:
        return []
    if isinstance(value, str):
        chunks: list[Any] = [value]
    elif isinstance(value, (list, tuple)):
        chunks = list(value)
    else:
        chunks = [value]

    items: list[str] = []
    for chunk in chunks:
        if chunk is None:
            continue
        for line in str(chunk).splitlines():
            text = line.strip()
            if not text:
                continue
            stripped = LIST_ITEM_PREFIX_PATTERN.sub("", text, count=1).strip()
            items.append(stripped or text)
    return items


def format_excel_multiline(items: Any) -> str:
    """Render list items for a single Excel cell: plain when single, numbered otherwise."""
    if items is None:
        values: list[str] = []
    elif isinstance(items, str):
        text = items.strip()
        values = [text] if text else []
    else:
        values = [
            str(item).strip()
            for item in items
            if item is not None and str(item).strip()
        ]
    if not values:
        return ""
    if len(values) == 1:
        return values[0]
    return "\n".join(f"{index}. {value}" for index, value in enumerate(values, start=1))


def _display_width(text: str) -> float:
    width = 0.0
    for char in text:
        width += 2.0 if unicodedata.east_asian_width(char) in ("W", "F") else 1.0
    return width


def _wrapped_line_count(text: Any, usable_width: int) -> int:
    value = "" if text is None else str(text)
    value = value.replace("\r\n", "\n").replace("\r", "\n")
    total = 0
    for physical_line in value.split("\n"):
        total += max(1, math.ceil(_display_width(physical_line) / usable_width))
    return max(1, total)


def estimate_row_height(
    values: Any,
    column_widths: list[int] | None = None,
) -> int:
    """Estimate an Excel row height from each column's wrapped line count.

    The tallest column wins (never the sum), so short cells in wide columns do
    not inflate the height. Result is clamped to ``[24, 409]`` points.
    """
    items = list(values or [])
    widths = list(column_widths or [])
    max_lines = 1
    for index, value in enumerate(items):
        column_width = int(widths[index]) if index < len(widths) else DEFAULT_COLUMN_WIDTH
        usable_width = max(column_width - 2, 8)
        max_lines = max(max_lines, _wrapped_line_count(value, usable_width))
    height = max_lines * ROW_HEIGHT_PER_LINE + ROW_HEIGHT_PADDING
    return min(max(MIN_ROW_HEIGHT, height), MAX_ROW_HEIGHT)


def escape_md_table_cell(text: Any) -> str:
    """Escape a Markdown table cell so literal pipes do not split the row."""
    if text is None:
        return ""
    return str(text).replace("|", "\\|")


def split_md_table_row(line: str) -> list[str]:
    """Split a Markdown table row, honouring ``\\|`` escapes and stripping ``|`` edges."""
    text = str(line).strip()
    if text.startswith("|"):
        text = text[1:]
    if text.endswith("|") and not text.endswith("\\|"):
        text = text[:-1]

    cells: list[str] = []
    buffer: list[str] = []
    index = 0
    while index < len(text):
        char = text[index]
        if char == "\\" and index + 1 < len(text) and text[index + 1] == "|":
            buffer.append("|")
            index += 2
            continue
        if char == "|":
            cells.append("".join(buffer).strip())
            buffer = []
            index += 1
            continue
        buffer.append(char)
        index += 1
    cells.append("".join(buffer).strip())
    return cells
