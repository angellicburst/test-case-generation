from __future__ import annotations

from copy import deepcopy
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from case_baseline import (  # noqa: E402
    _merge_duplicate_entries,
    assign_missing_ids,
    canonical_tree_order,
    merge_artifacts_to_ir,
    normalize_baseline,
    parse_excel_to_ir,
    parse_markdown_to_ir,
    renumber_cases,
    validate_baseline,
)
from create_excel_cases import (  # noqa: E402
    build_workbook,
    build_workbook_from_ir,
    save_workbook_atomic,
)
from render_markdown_cases import render_markdown  # noqa: E402
from traceability_matrix import (  # noqa: E402
    describe_traceability_mismatch,
    format_trace_id_list,
    load_traceability_matrix,
    parse_trace_id_list,
    sync_traceability_from_cases,
    validate_traceability_entries,
)
from sync_traceability import _copy_to_temp, sync_atomically, sync_excel, sync_markdown  # noqa: E402
from validate_excel_cases import validate_excel_file  # noqa: E402
from validate_markdown_lists import fix, validate  # noqa: E402
from validate_traceability_bundle import validate_bundle  # noqa: E402
from validation_utils import (  # noqa: E402
    FileLockedError,
    escape_md_table_cell,
    estimate_row_height,
    format_excel_multiline,
    is_valid_date_or_placeholder,
    normalize_list_items,
    split_md_table_row,
    write_text_atomic,
)


VALID_MARKDOWN = """# 测试用例 - 2026-11-17 - [注册]
> 用例数量: 用例 1 条

- [ ] 注册模块
  - [ ] 功能测试
    - [ ] 正向-提交合法注册-P0
      > [需求追溯] 来源：注册需求.md｜定位：2.1 手机号规则｜规则：手机号必须为11位数字
      - 前置条件
        - 用户未注册。
      - 测试步骤
        - 提交注册。
      - 预期结果
        - 注册成功。
"""

GROUPED_MARKDOWN = """# 测试用例 - 2026-11-17 - [注册]
> 用例数量: 用例 1 条

- [ ] 注册模块
  - [ ] 功能测试
    - [ ] 分组
      - [ ] 正向-提交合法注册-P0
        > [需求追溯] 来源：注册需求.md｜定位：2.1 手机号规则｜规则：手机号必须为11位数字
        - 前置条件
          - 用户未注册。
        - 测试步骤
          - 提交注册。
        - 预期结果
          - 注册成功。
"""

TRACEABILITY_MATRIX = {
    "TRACE-REG-001": {
        "id": "TRACE-REG-001",
        "source": "注册需求.md",
        "location": "2.1 手机号规则",
        "rule": "手机号必须为11位数字",
        "status": "已覆盖",
        "case_ids": ["TC-REG-001"],
        "case_titles": ["提交合法注册"],
    }
}


def excel_case(**overrides):
    case = {
        "requirement_id": "",
        "id": "TC-REG-001",
        "title": "提交合法注册",
        "type": "功能测试",
        "design_method": "等价类划分",
        "priority": "P0",
        "preconditions": "用户未注册",
        "test_data": "<独立测试账号>",
        "steps": "提交注册",
        "expected_results": "注册成功",
        "traceability_refs": ["TRACE-REG-001"],
        "notes": "无",
    }
    case.update(overrides)
    return case


def matrix_for_cases(*cases):
    return {
        "TRACE-REG-001": {
            **TRACEABILITY_MATRIX["TRACE-REG-001"],
            "case_ids": [str(case.get("id", case.get("case_id", ""))) for case in cases],
            "case_titles": [str(case.get("title", case.get("case_title", ""))) for case in cases],
        }
    }


def valid_meta(**overrides):
    meta = {"requirement_name": "注册需求", "release_date": "未指定日期"}
    meta.update(overrides)
    return meta


def valid_ir():
    return {
        "schema_version": "2.0",
        "meta": {
            "title": "测试用例 - 2026-11-24 - [注册]",
            "requirement_name": "注册",
            "release_date": "2026-11-24",
            "coverage": "覆盖注册主流程",
            "notes": "基于需求文档生成",
            "sheet_name": "测试用例",
        },
        "traceability": [
            {
                "id": "TRACE-REG-001",
                "source": "注册需求.md",
                "location": "1.1 手机号",
                "rule": "手机号必须为11位数字",
                "status": "已覆盖",
                "case_ids": ["TC-ZC-001"],
            }
        ],
        "cases": [
            {
                "id": "TC-ZC-001",
                "module": "注册模块",
                "type": "功能测试",
                "group": "",
                "polarity": "正向",
                "title": "[注册模块]_[注册]_[合法手机号]_[成功]",
                "design_method": "等价类划分",
                "priority": "P0",
                "pre_operations": ["用户未注册。"],
                "test_data": ["手机号: <测试手机号>"],
                "steps": ["提交合法手机号。"],
                "expected_results": ["注册成功。"],
                "post_operations": ["删除本次注册的账号。"],
                "traceability_refs": ["TRACE-REG-001"],
            }
        ],
    }


def build_valid_workbook(title, cases, matrix, sheet_name="测试用例", **meta_overrides):
    return build_workbook(
        title,
        cases,
        matrix,
        sheet_name=sheet_name,
        meta=valid_meta(**meta_overrides),
    )


def multi_traceability_matrix():
    return {
        **deepcopy(TRACEABILITY_MATRIX),
        "TRACE-REG-002": {
            "id": "TRACE-REG-002",
            "source": "注册需求.md",
            "location": "2.2 注册成功",
            "rule": "注册成功后显示新注册信息",
            "status": "已覆盖",
            "case_ids": ["TC-REG-001"],
            "case_titles": ["提交合法注册"],
        },
    }


def write_matrix(path: Path, matrix: dict[str, dict[str, object]]) -> None:
    path.write_text(
        json.dumps({"entries": list(matrix.values())}, ensure_ascii=False),
        encoding="utf-8",
    )


def subprocess_env():
    return {**os.environ, "PYTHONIOENCODING": "utf-8"}


def subprocess_env_without_encoding_override():
    environment = os.environ.copy()
    environment.pop("PYTHONIOENCODING", None)
    return environment


class MarkdownValidationTests(unittest.TestCase):
    def test_valid_document(self):
        self.assertEqual(validate(VALID_MARKDOWN, TRACEABILITY_MATRIX), [])

    def test_missing_release_date_uses_documented_placeholder(self):
        text = VALID_MARKDOWN.replace("2026-11-17", "未指定日期")
        self.assertEqual(validate(text, TRACEABILITY_MATRIX), [])

    def test_invalid_calendar_date_rejected(self):
        invalid_day = VALID_MARKDOWN.replace("2026-11-17", "2026-02-29")
        self.assertTrue(
            any("不是有效日历日期" in error for error in validate(invalid_day, TRACEABILITY_MATRIX))
        )
        invalid_month = VALID_MARKDOWN.replace("2026-11-17", "2026-19-99")
        self.assertTrue(
            any("不是有效日历日期" in error for error in validate(invalid_month, TRACEABILITY_MATRIX))
        )
        leap_day = VALID_MARKDOWN.replace("2026-11-17", "2024-02-29")
        self.assertEqual(validate(leap_day, TRACEABILITY_MATRIX), [])

    def test_extra_direct_field_rejected(self):
        text = VALID_MARKDOWN.replace(
            "      - 前置条件\n",
            "      - 自定义字段\n        - 任意内容。\n      - 前置条件\n",
        )
        errors = validate(text, TRACEABILITY_MATRIX)
        self.assertTrue(any("发现非法字段“自定义字段”" in error for error in errors))

        nested = VALID_MARKDOWN.replace(
            "        - 提交注册。\n",
            "        - 打开页面。\n          - 在独立窗口打开。\n",
        )
        self.assertEqual(validate(nested, TRACEABILITY_MATRIX), [])

    def test_extra_direct_text_rejected(self):
        text = VALID_MARKDOWN.replace(
            "      > [需求追溯] 来源：注册需求.md｜定位：2.1 手机号规则｜规则：手机号必须为11位数字\n",
            "      > [需求追溯] 来源：注册需求.md｜定位：2.1 手机号规则｜规则：手机号必须为11位数字\n"
            "      说明：这是未定义的直接内容。\n",
        )
        errors = validate(text, TRACEABILITY_MATRIX)
        self.assertTrue(
            any("发现非法直接内容“说明：这是未定义的直接内容。”" in error for error in errors)
        )

    def test_invalid_test_type_rejected(self):
        text = VALID_MARKDOWN.replace("  - [ ] 功能测试", "  - [ ] 自定义类型")
        self.assertTrue(
            any("测试类型非法" in error for error in validate(text, TRACEABILITY_MATRIX))
        )
        for test_type in ("回归测试", "兼容性测试", "可靠性测试"):
            self.assertEqual(
                validate(
                    VALID_MARKDOWN.replace("功能测试", test_type),
                    TRACEABILITY_MATRIX,
                ),
                [],
            )

    def test_empty_document_rejected(self):
        errors = validate("# 测试用例\n> 用例数量: 用例 0 条\n", TRACEABILITY_MATRIX)
        self.assertTrue(any("未包含任何有效用例任务项" in error for error in errors))

    def test_missing_required_fields_rejected(self):
        text = VALID_MARKDOWN.replace(
            "      > [需求追溯] 来源：注册需求.md｜定位：2.1 手机号规则｜规则：手机号必须为11位数字\n",
            "",
        )
        errors = validate(text, TRACEABILITY_MATRIX)
        self.assertTrue(any("标题下缺少需求追溯引用" in error for error in errors))
        text = VALID_MARKDOWN.replace("      - 前置条件\n        - 用户未注册。\n", "")
        text = text.replace("      - 预期结果\n        - 注册成功。\n", "")
        errors = validate(text, TRACEABILITY_MATRIX)
        self.assertTrue(any("缺少必填字段“前置条件”" in error for error in errors))
        self.assertTrue(any("缺少必填字段“预期结果”" in error for error in errors))

    def test_duplicate_field_rejected(self):
        text = VALID_MARKDOWN.replace(
            "      - 预期结果\n        - 注册成功。\n",
            "      - 预期结果\n        - 注册成功。\n      - 预期结果\n        - 列表可见。\n",
        )
        self.assertTrue(any("重复的“预期结果”字段" in error for error in validate(text, TRACEABILITY_MATRIX)))

    def test_traceability_must_be_first_and_nonempty(self):
        with_blank_line = VALID_MARKDOWN.replace(
            "    - [ ] 正向-提交合法注册-P0\n      >",
            "    - [ ] 正向-提交合法注册-P0\n\n      >",
        )
        self.assertEqual(validate(with_blank_line, TRACEABILITY_MATRIX), [])
        misplaced = VALID_MARKDOWN.replace(
            "      > [需求追溯] 来源：注册需求.md｜定位：2.1 手机号规则｜规则：手机号必须为11位数字\n",
            "      - 前置条件\n        - 用户未注册。\n      > [需求追溯] 来源：注册需求.md｜定位：2.1 手机号规则｜规则：手机号必须为11位数字\n",
        )
        self.assertTrue(any("标题下首个非空行填写需求追溯" in error for error in validate(misplaced, TRACEABILITY_MATRIX)))
        empty = VALID_MARKDOWN.replace(
            "      > [需求追溯] 来源：注册需求.md｜定位：2.1 手机号规则｜规则：手机号必须为11位数字\n",
            "      >    \n",
        )
        self.assertTrue(any("标题下缺少需求追溯引用" in error for error in validate(empty, TRACEABILITY_MATRIX)))

    def test_traceability_requires_source_location_and_rule(self):
        text = VALID_MARKDOWN.replace(
            "[需求追溯] 来源：注册需求.md｜定位：2.1 手机号规则｜规则：手机号必须为11位数字",
            "[需求追溯] REQ-001",
        )
        self.assertTrue(any("不得仅使用内部编号" in error for error in validate(text, TRACEABILITY_MATRIX)))
        legacy = VALID_MARKDOWN.replace("[需求追溯] 来源：", "需求追溯：来源：")
        self.assertTrue(any("需求追溯必须填写" in error for error in validate(legacy, TRACEABILITY_MATRIX)))
        missing_space = VALID_MARKDOWN.replace("[需求追溯] 来源：", "[需求追溯]来源：")
        self.assertTrue(any("需求追溯必须填写" in error for error in validate(missing_space, TRACEABILITY_MATRIX)))

    def test_count_uses_case_items(self):
        text = VALID_MARKDOWN.replace("用例 1 条", "用例 2 条")
        self.assertTrue(any("实际用例任务项数量为 1" in error for error in validate(text, TRACEABILITY_MATRIX)))

    def test_markdown_title_declaration_and_indentation_required(self):
        invalid_title = VALID_MARKDOWN.replace("# 测试用例 - 2026-11-17 - [注册]", "# 注册用例")
        self.assertTrue(
            any("文档首行必须" in error for error in validate(invalid_title, TRACEABILITY_MATRIX))
        )
        moved_declaration = VALID_MARKDOWN.replace(
            "> 用例数量: 用例 1 条\n\n",
            "\n说明\n> 用例数量: 用例 1 条\n",
        )
        self.assertTrue(
            any("标题后首个非空行" in error for error in validate(moved_declaration, TRACEABILITY_MATRIX))
        )
        invalid_indent = VALID_MARKDOWN.replace("  - [ ] 功能测试", "   - [ ] 功能测试")
        self.assertTrue(
            any("缩进必须为 2 个空格" in error for error in validate(invalid_indent, TRACEABILITY_MATRIX))
        )

    def test_orphan_fields_do_not_count_as_cases(self):
        text = "# 测试用例\n> 用例数量: 用例 0 条\n- [ ] 模块\n  - [ ] 功能测试\n    - 预期结果\n      - 孤立字段。\n"
        self.assertTrue(any("未包含任何有效用例任务项" in error for error in validate(text, TRACEABILITY_MATRIX)))

    def test_invalid_type_checked_case_and_duplicate_title_rejected(self):
        text = VALID_MARKDOWN.replace("正向-提交合法注册", "边界-提交合法注册")
        self.assertTrue(any("用例类型非法" in error for error in validate(text, TRACEABILITY_MATRIX)))
        self.assertTrue(any("不得使用已勾选状态" in error for error in validate(VALID_MARKDOWN.replace("- [ ] 正向", "- [x] 正向"), TRACEABILITY_MATRIX)))
        duplicate = VALID_MARKDOWN.replace("用例 1 条", "用例 2 条") + VALID_MARKDOWN.split("- [ ] 注册模块\n", 1)[1].replace("注册模块", "其他模块")
        self.assertTrue(any("用例标题“提交合法注册”重复" in error for error in validate(duplicate, TRACEABILITY_MATRIX)))

    def test_list_format_and_fix(self):
        text = VALID_MARKDOWN.replace("        - 提交注册。\n", "        - 打开页面。\n        - 提交注册。\n")
        self.assertTrue(any("必须全部使用有序列表" in error for error in validate(text, TRACEABILITY_MATRIX)))
        self.assertEqual(validate(fix(text), TRACEABILITY_MATRIX), [])

    def test_fix_only_rewrites_direct_field_items(self):
        text = VALID_MARKDOWN.replace(
            "        - 提交注册。\n",
            "        - 打开页面。\n          - 在独立窗口打开。\n        - 提交注册。\n",
        )
        fixed = fix(text)
        self.assertIn("        1. 打开页面。\n          - 在独立窗口打开。\n        2. 提交注册。\n", fixed)

    def test_markdown_cli_and_utf8(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "cases.md"
            matrix = Path(directory) / "matrix.json"
            write_matrix(matrix, TRACEABILITY_MATRIX)
            target.write_text(VALID_MARKDOWN, encoding="utf-8")
            result = subprocess.run(
                [sys.executable, str(ROOT / "scripts" / "validate_markdown_lists.py"), str(target), "--traceability-matrix", str(matrix)],
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=subprocess_env(),
                check=False,
            )
            self.assertEqual(result.returncode, 0, result.stderr)

            target.write_text(VALID_MARKDOWN, encoding="gbk")
            result = subprocess.run(
                [sys.executable, str(ROOT / "scripts" / "validate_markdown_lists.py"), str(target)],
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=subprocess_env(),
                check=False,
            )
            self.assertNotEqual(result.returncode, 0)

    def test_cli_errors_are_utf8_without_environment_override(self):
        with tempfile.TemporaryDirectory() as directory:
            matrix = Path(directory) / "matrix.json"
            input_json = Path(directory) / "cases.json"
            missing_markdown = Path(directory) / "missing.md"
            missing_excel = Path(directory) / "missing.xlsx"
            write_matrix(matrix, TRACEABILITY_MATRIX)
            input_json.write_text(
                json.dumps({"meta": valid_meta(), "cases": [excel_case()]}, ensure_ascii=False),
                encoding="utf-8",
            )
            commands = [
                (
                    [
                        sys.executable,
                        str(ROOT / "scripts" / "validate_markdown_lists.py"),
                        str(missing_markdown),
                        "--traceability-matrix",
                        str(matrix),
                    ],
                    "文件不存在",
                ),
                (
                    [
                        sys.executable,
                        str(ROOT / "scripts" / "validate_excel_cases.py"),
                        str(missing_excel),
                        "--traceability-matrix",
                        str(matrix),
                    ],
                    "文件不存在",
                ),
                (
                    [
                        sys.executable,
                        str(ROOT / "scripts" / "validate_traceability_bundle.py"),
                        "--traceability-matrix",
                        str(matrix),
                        "--markdown",
                        str(missing_markdown),
                        "--excel",
                        str(missing_excel),
                        "--coverage-only",
                    ],
                    "文件不存在",
                ),
                (
                    [
                        sys.executable,
                        str(ROOT / "scripts" / "sync_traceability.py"),
                        "--traceability-matrix",
                        str(matrix),
                        "--markdown",
                        str(missing_markdown),
                    ],
                    "文件不存在",
                ),
                (
                    [
                        sys.executable,
                        str(ROOT / "scripts" / "create_excel_cases.py"),
                        "--output",
                        str(directory),
                        "--input-json",
                        str(input_json),
                        "--traceability-matrix",
                        str(matrix),
                    ],
                    "输出路径必须是文件路径",
                ),
            ]
            for command, expected_error in commands:
                with self.subTest(command=Path(command[1]).name):
                    result = subprocess.run(
                        command,
                        capture_output=True,
                        env=subprocess_env_without_encoding_override(),
                        check=False,
                    )
                    self.assertNotEqual(result.returncode, 0)
                    self.assertIn(expected_error, result.stderr.decode("utf-8"))

    def test_markdown_cli_fix_writes_without_temp_leftovers(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "cases.md"
            matrix = Path(directory) / "matrix.json"
            write_matrix(matrix, TRACEABILITY_MATRIX)
            target.write_text(
                VALID_MARKDOWN.replace(
                    "        - 提交注册。\n",
                    "        - 打开页面。\n        - 提交注册。\n",
                ),
                encoding="utf-8",
            )
            result = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "validate_markdown_lists.py"),
                    str(target),
                    "--traceability-matrix",
                    str(matrix),
                    "--fix",
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=subprocess_env(),
                check=False,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            fixed = target.read_text(encoding="utf-8")
            self.assertIn("        1. 打开页面。\n        2. 提交注册。\n", fixed)
            self.assertEqual(list(Path(directory).glob(".*.tmp")), [])

    def test_markdown_fix_preserves_line_endings(self):
        source = VALID_MARKDOWN.replace(
            "        - 提交注册。\n",
            "        - 打开页面。\n        - 提交注册。\n",
        )
        for newline, label in (("\n", "LF"), ("\r\n", "CRLF")):
            with self.subTest(label=label):
                with tempfile.TemporaryDirectory() as directory:
                    target = Path(directory) / "cases.md"
                    matrix = Path(directory) / "matrix.json"
                    write_matrix(matrix, TRACEABILITY_MATRIX)
                    target.write_bytes(source.replace("\n", newline).encode("utf-8"))
                    result = subprocess.run(
                        [
                            sys.executable,
                            str(ROOT / "scripts" / "validate_markdown_lists.py"),
                            str(target),
                            "--traceability-matrix",
                            str(matrix),
                            "--fix",
                        ],
                        capture_output=True,
                        text=True,
                        encoding="utf-8",
                        env=subprocess_env(),
                        check=False,
                    )
                    self.assertEqual(result.returncode, 0, result.stderr)
                    data = target.read_bytes()
                    self.assertIn(
                        f"        1. 打开页面。{newline}".encode("utf-8"),
                        data,
                    )
                    if newline == "\r\n":
                        self.assertEqual(data.count(b"\n"), data.count(b"\r\n"))
                    else:
                        self.assertNotIn(b"\r", data)

    def test_traceability_must_match_matrix(self):
        altered = VALID_MARKDOWN.replace("手机号必须为11位数字", "任意规则")
        errors = validate(altered, TRACEABILITY_MATRIX)
        self.assertTrue(any("与矩阵记录不一致" in error for error in errors))
        self.assertTrue(any("未提供需求追溯矩阵" in error for error in validate(VALID_MARKDOWN)))

    def test_markdown_supports_multiple_traceability_entries(self):
        matrix = multi_traceability_matrix()
        text = VALID_MARKDOWN.replace(
            "      > [需求追溯] 来源：注册需求.md｜定位：2.1 手机号规则｜规则：手机号必须为11位数字\n",
            "      > [需求追溯] 来源：注册需求.md｜定位：2.1 手机号规则｜规则：手机号必须为11位数字\n"
            "      > [需求追溯] 来源：注册需求.md｜定位：2.2 注册成功｜规则：注册成功后显示新注册信息\n",
        )
        self.assertEqual(validate(text, matrix), [])
        self.assertTrue(
            any(
                "需求追溯与矩阵记录不一致" in error
                for error in validate(VALID_MARKDOWN, matrix)
            )
        )

    def test_markdown_rejects_stale_matrix_title(self):
        matrix = {
            "TRACE-REG-001": {
                **TRACEABILITY_MATRIX["TRACE-REG-001"],
                "case_titles": ["已删除的注册用例"],
            }
        }
        errors = validate(VALID_MARKDOWN, matrix)
        self.assertTrue(
            any("未出现在文档中" in error for error in errors)
        )

    def test_assumption_case_marker_does_not_match_matrix_title(self):
        marked = VALID_MARKDOWN.replace("正向-提交合法注册-P0", "正向-提交合法注册[待确认]-P0")
        errors = validate(marked, TRACEABILITY_MATRIX)
        self.assertTrue(any("不得包含旧版“[待确认]”标记" in error for error in errors))

    def test_covered_entry_requires_markdown_case_titles(self):
        matrix = {
            "TRACE-REG-001": {
                **TRACEABILITY_MATRIX["TRACE-REG-001"],
                "case_titles": [],
            }
        }
        errors = validate(VALID_MARKDOWN, matrix)
        self.assertTrue(
            any("缺少 Markdown 用例标题关联" in error for error in errors),
            errors,
        )

    def test_grouped_invalid_case_reports_single_root_cause(self):
        self.assertEqual(validate(GROUPED_MARKDOWN, TRACEABILITY_MATRIX), [])
        variants = {
            "用例类型非法": GROUPED_MARKDOWN.replace(
                "正向-提交合法注册", "边界-提交合法注册"
            ),
            "不得使用已勾选状态": GROUPED_MARKDOWN.replace("- [ ] 正向", "- [x] 正向"),
            "不得包含用例编号": GROUPED_MARKDOWN.replace(
                "正向-提交合法注册", "正向-TC-001-提交合法注册"
            ),
        }
        for expected, text in variants.items():
            with self.subTest(expected=expected):
                errors = validate(text, TRACEABILITY_MATRIX)
                self.assertTrue(any(expected in error for error in errors), errors)
                self.assertFalse(
                    any("6 空格层级只能用于用例任务项" in error for error in errors),
                    errors,
                )
        illegal_group = GROUPED_MARKDOWN.replace(
            "    - [ ] 分组\n", "    - [ ] 分组\n      - [ ] 非法子分组\n"
        )
        self.assertTrue(
            any(
                "6 空格层级只能用于用例任务项" in error
                for error in validate(illegal_group, TRACEABILITY_MATRIX)
            )
        )

    def test_invalid_priority_reported_without_level_error(self):
        cases = {
            4: VALID_MARKDOWN.replace("正向-提交合法注册-P0", "正向-提交合法注册-P5"),
            6: GROUPED_MARKDOWN.replace("正向-提交合法注册-P0", "正向-提交合法注册-P5"),
        }
        for indent, text in cases.items():
            with self.subTest(indent=indent):
                errors = validate(text, TRACEABILITY_MATRIX)
                self.assertTrue(any("用例优先级非法" in error for error in errors), errors)
                self.assertFalse(
                    any("6 空格层级只能用于用例任务项" in error for error in errors),
                    errors,
                )

    def test_pending_marker_reported_at_its_own_line(self):
        text = VALID_MARKDOWN + "- [ ] 其他模块[待确认]\n"
        expected_line = next(
            index
            for index, line in enumerate(text.splitlines(), start=1)
            if "[待确认]" in line
        )
        errors = validate(text, TRACEABILITY_MATRIX)
        self.assertEqual(
            [error for error in errors if "不得包含旧版" in error],
            [f"第 {expected_line} 行不得包含旧版“[待确认]”标记。"],
        )

    def test_markdown_has_no_cleanup_structure_requirement(self):
        self.assertEqual(validate(VALID_MARKDOWN, TRACEABILITY_MATRIX), [])

    def test_invalid_traceability_matrix_reports_single_root_cause(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "cases.md"
            matrix = Path(directory) / "invalid-matrix.json"
            target.write_text(VALID_MARKDOWN, encoding="utf-8")
            matrix.write_text('{"entries": []}', encoding="utf-8")
            result = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "validate_markdown_lists.py"),
                    str(target),
                    "--traceability-matrix",
                    str(matrix),
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=subprocess_env(),
                check=False,
            )
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("entries 数组", result.stderr)
            self.assertNotIn("未提供需求追溯矩阵", result.stderr)

    def test_all_task_items_must_use_unchecked_state(self):
        for original, replacement in (
            ("- [ ] 注册模块", "- [x] 注册模块"),
            ("- [ ] 功能测试", "- [X] 功能测试"),
            ("- [ ] 正向-提交合法注册-P0", "- [ x] 正向-提交合法注册-P0"),
            ("- [ ] 正向-提交合法注册-P0", "- [X]正向-提交合法注册-P0"),
        ):
            with self.subTest(replacement=replacement):
                text = VALID_MARKDOWN.replace(original, replacement)
                errors = validate(text, TRACEABILITY_MATRIX)
                self.assertTrue(
                    any("任务项状态必须为 `- [ ]`" in error for error in errors),
                    errors,
                )

    def test_case_field_order_must_be_standard(self):
        reordered = VALID_MARKDOWN.replace(
            "      - 前置条件\n"
            "        - 用户未注册。\n"
            "      - 测试步骤\n"
            "        - 提交注册。\n",
            "      - 测试步骤\n"
            "        - 提交注册。\n"
            "      - 前置条件\n"
            "        - 用户未注册。\n",
        )
        self.assertTrue(
            any("字段顺序必须为" in error for error in validate(reordered, TRACEABILITY_MATRIX))
        )
        with_blank_lines = VALID_MARKDOWN.replace(
            "      - 测试步骤\n",
            "\n      - 测试步骤\n",
        )
        self.assertEqual(validate(with_blank_lines, TRACEABILITY_MATRIX), [])

    def test_compact_title_date_rejected(self):
        compact = VALID_MARKDOWN.replace("2026-11-17", "20261117")
        self.assertTrue(
            any("文档首行必须" in error for error in validate(compact, TRACEABILITY_MATRIX))
        )


class TraceabilityMatrixValidationTests(unittest.TestCase):
    def test_matrix_accepts_utf8_bom(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "matrix.json"
            payload = json.dumps(
                {"entries": list(TRACEABILITY_MATRIX.values())},
                ensure_ascii=False,
            )
            path.write_bytes(b"\xef\xbb\xbf" + payload.encode("utf-8"))
            matrix, errors = load_traceability_matrix(path)
            self.assertEqual(errors, [])
            self.assertIn("TRACE-REG-001", matrix)

    def test_matrix_rejects_multiline_fields(self):
        for field in ("source", "location", "rule"):
            with self.subTest(field=field):
                entry = {
                    **TRACEABILITY_MATRIX["TRACE-REG-001"],
                    field: "第一行\n第二行",
                }
                with tempfile.TemporaryDirectory() as directory:
                    path = Path(directory) / "matrix.json"
                    write_matrix(path, {entry["id"]: entry})
                    _, errors = load_traceability_matrix(path)
                    self.assertTrue(
                        any("不得包含换行" in error for error in errors),
                        errors,
                    )


class TraceabilityAndBaselineTests(unittest.TestCase):
    COVERED = {
        "id": "TRACE-A",
        "source": "需求.md",
        "location": "1.1 规则",
        "rule": "规则 A",
        "status": "已覆盖",
        "case_ids": ["TC-A-001"],
    }
    PENDING = {
        "id": "TRACE-B",
        "source": "需求.md",
        "location": "1.2 规则",
        "status": "待确认",
        "issue": "未明确项",
        "assumption": "当前假设",
        "question": "确认问题？",
    }

    def test_validate_entries_accepts_covered_and_pending(self):
        matrix, errors = validate_traceability_entries(
            [deepcopy(self.COVERED), deepcopy(self.PENDING)]
        )
        self.assertEqual(errors, [])
        self.assertEqual(list(matrix), ["TRACE-A", "TRACE-B"])
        self.assertEqual(matrix["TRACE-A"]["case_ids"], ["TC-A-001"])
        self.assertNotIn("rule", matrix["TRACE-B"])

    def test_validate_entries_rejects_legacy_case_titles(self):
        entry = {**deepcopy(self.COVERED), "case_titles": ["提交合法注册"]}
        _, errors = validate_traceability_entries([entry])
        self.assertTrue(any("case_titles" in error for error in errors), errors)

    def test_validate_entries_rejects_invalid_id_characters(self):
        entry = {**deepcopy(self.COVERED), "id": "TRACE A[1]"}
        _, errors = validate_traceability_entries([entry])
        self.assertTrue(any("非法字符" in error for error in errors), errors)

    def test_validate_entries_rejects_duplicate_rule_triple(self):
        duplicate = {**deepcopy(self.COVERED), "id": "TRACE-A2"}
        _, errors = validate_traceability_entries([deepcopy(self.COVERED), duplicate])
        self.assertTrue(any("重复" in error for error in errors), errors)

    def test_validate_entries_rejects_pending_fields_on_covered(self):
        entry = {**deepcopy(self.COVERED), "issue": "不该出现"}
        _, errors = validate_traceability_entries([entry])
        self.assertTrue(any("不得填写待确认字段" in error for error in errors), errors)

    def test_validate_entries_requires_case_ids_for_covered(self):
        entry = {**deepcopy(self.COVERED), "case_ids": []}
        _, errors = validate_traceability_entries([entry])
        self.assertTrue(any("必须至少关联一个用例" in error for error in errors), errors)

    def test_validate_entries_rejects_case_ids_on_pending(self):
        entry = {**deepcopy(self.PENDING), "case_ids": ["TC-A-001"]}
        _, errors = validate_traceability_entries([entry])
        self.assertTrue(any("不得关联用例" in error for error in errors), errors)

    def test_parse_and_format_trace_id_list(self):
        ids, errors = parse_trace_id_list("TRACE-A, TRACE-B，TRACE-B")
        self.assertEqual(ids, ["TRACE-A", "TRACE-B"])
        self.assertTrue(any("重复" in error for error in errors), errors)
        self.assertEqual(parse_trace_id_list("")[0], [])
        self.assertEqual(
            format_trace_id_list([{"id": "TRACE-A"}, "TRACE-B"]),
            "TRACE-A, TRACE-B",
        )

    def test_describe_mismatch_covers_all_differences(self):
        self.assertEqual(describe_traceability_mismatch(["A", "B"], ["A", "B"]), "")
        self.assertIn(
            "顺序不一致", describe_traceability_mismatch(["A", "B"], ["B", "A"])
        )
        self.assertIn("缺失: B", describe_traceability_mismatch(["A", "B"], ["A"]))
        self.assertIn("多余: C", describe_traceability_mismatch(["A"], ["A", "C"]))

    def test_sync_upgrades_pending_coverage_in_case_order(self):
        traceability = [
            {
                "id": "TRACE-A",
                "source": "需求.md",
                "location": "1.1",
                "rule": "规则 A",
                "status": "待覆盖",
            }
        ]
        cases = [
            {"id": "TC-A-001", "traceability_refs": ["TRACE-A"]},
            {"id": "TC-A-002", "traceability_refs": ["TRACE-A"]},
        ]
        result, errors, warnings = sync_traceability_from_cases(traceability, cases)
        self.assertEqual(errors, [])
        self.assertEqual(warnings, [])
        self.assertEqual(result["TRACE-A"]["status"], "已覆盖")
        self.assertEqual(result["TRACE-A"]["case_ids"], ["TC-A-001", "TC-A-002"])
        self.assertEqual(traceability[0]["status"], "待覆盖")

    def test_sync_accepts_matrix_dict_input(self):
        matrix = {
            "TRACE-A": {
                "source": "需求.md",
                "location": "1.1",
                "rule": "规则 A",
                "status": "待覆盖",
            }
        }
        cases = [{"id": "TC-A-001", "traceability_refs": ["TRACE-A"]}]
        result, errors, warnings = sync_traceability_from_cases(matrix, cases)
        self.assertEqual(errors, [])
        self.assertEqual(warnings, [])
        self.assertEqual(result["TRACE-A"]["status"], "已覆盖")

    def test_sync_strips_status_whitespace_before_settling(self):
        pending = {**deepcopy(self.PENDING), "status": " 待确认 "}
        _, errors, _ = sync_traceability_from_cases(
            [pending], [{"id": "TC-A-001", "traceability_refs": ["TRACE-B"]}]
        )
        self.assertTrue(any("不得引用待确认条目" in error for error in errors), errors)

        covering = {
            "id": "TRACE-A",
            "source": "需求.md",
            "location": "1.1",
            "rule": "规则 A",
            "status": " 待覆盖 ",
        }
        result, errors, _ = sync_traceability_from_cases(
            [covering], [{"id": "TC-A-001", "traceability_refs": ["TRACE-A"]}]
        )
        self.assertEqual(errors, [])
        self.assertEqual(result["TRACE-A"]["status"], "已覆盖")

    def test_describe_mismatch_reports_duplicates_not_order(self):
        message = describe_traceability_mismatch(["A", "B"], ["A", "B", "B"])
        self.assertNotIn("顺序不一致", message)
        self.assertIn("重复项", message)

    def test_sync_blocks_reference_to_pending_entry(self):
        cases = [{"id": "TC-A-001", "traceability_refs": ["TRACE-B"]}]
        _, errors, _ = sync_traceability_from_cases(
            [deepcopy(self.PENDING)], cases
        )
        self.assertTrue(any("不得引用待确认条目" in error for error in errors), errors)

    def test_sync_downgrades_coverage_without_cases(self):
        result, errors, warnings = sync_traceability_from_cases(
            [deepcopy(self.COVERED)], []
        )
        self.assertEqual(errors, [])
        self.assertEqual(result["TRACE-A"]["status"], "待覆盖")
        self.assertTrue(any("回退为待覆盖" in warning for warning in warnings), warnings)


class ExcelTests(unittest.TestCase):
    def test_matrix_status_is_chinese_and_controls_case_links(self):
        pending = {
            "id": "TRACE-REG-002",
            "source": "注册需求.md",
            "location": "2.2 未定义规则",
            "rule": "需求未明确预期",
            "status": "待确认",
            "issue": "异常提示未定义",
            "assumption": "仅记录问题，不生成用例",
            "question": "请确认预期提示",
        }
        with tempfile.TemporaryDirectory() as directory:
            matrix = Path(directory) / "matrix.json"
            write_matrix(matrix, {pending["id"]: pending})
            loaded, errors = load_traceability_matrix(matrix)
            self.assertEqual(errors, [])
            self.assertEqual(loaded["TRACE-REG-002"]["status"], "待确认")
            cases = [excel_case(traceability_refs=["TRACE-REG-002"])]
            with self.assertRaisesRegex(ValueError, "只能关联已覆盖"):
                build_workbook("测试用例", cases, loaded, meta=valid_meta())
        invalid_status = {**pending, "status": "pending", "case_ids": [], "case_titles": []}
        with tempfile.TemporaryDirectory() as directory:
            matrix = Path(directory) / "matrix.json"
            write_matrix(matrix, {invalid_status["id"]: invalid_status})
            _, errors = load_traceability_matrix(matrix)
            self.assertTrue(any("status 必须为“待覆盖”“已覆盖”或“待确认”" in error for error in errors))

    def test_draft_matrix_cannot_pass_final_format_validation(self):
        draft = {
            "TRACE-REG-001": {
                **TRACEABILITY_MATRIX["TRACE-REG-001"],
                "status": "待覆盖",
                "case_ids": [],
                "case_titles": [],
            }
        }
        self.assertTrue(
            any("存在待覆盖条目" in error for error in validate(VALID_MARKDOWN, draft))
        )
        self.assertFalse(any("未在需求追溯矩阵中关联" in error for error in validate(VALID_MARKDOWN, draft)))
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "draft.xlsx"
            build_valid_workbook("测试用例", [excel_case()], TRACEABILITY_MATRIX).save(target)
            errors = validate_excel_file(target, draft)
            self.assertTrue(any("存在待覆盖条目" in error for error in errors))
            self.assertFalse(any("未在需求追溯矩阵中关联" in error for error in errors))

    def test_matrix_separates_case_id_and_title_namespaces(self):
        entry = {
            **TRACEABILITY_MATRIX["TRACE-REG-001"],
            "case_ids": ["TC-SHARED-001"],
            "case_titles": ["TC-SHARED-001"],
        }
        with tempfile.TemporaryDirectory() as directory:
            matrix = Path(directory) / "matrix.json"
            write_matrix(matrix, {entry["id"]: entry})
            _, errors = load_traceability_matrix(matrix)
            self.assertEqual(errors, [])

    def test_build_and_validate_default_workbook(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "cases.xlsx"
            cases = [excel_case()]
            matrix = matrix_for_cases(*cases)
            build_valid_workbook("测试用例", cases, matrix).save(target)
            self.assertEqual(validate_excel_file(target, matrix), [])
            workbook = load_workbook(target)
            self.assertEqual(workbook.sheetnames, ["测试用例", "生成说明"])
            self.assertEqual(workbook["生成说明"]["A1"].value, "测试用例生成说明")
            self.assertEqual(workbook["生成说明"]["B4"].value, "1")
            self.assertEqual(
                workbook["测试用例"]["K3"].value,
                "[需求追溯] 来源：注册需求.md｜定位：2.1 手机号规则｜规则：手机号必须为11位数字\n无",
            )
            self.assertIsNone(workbook["测试用例"]["A3"].value)

    def test_excel_traceability_uses_explicit_source_and_preserves_notes(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "traceability.xlsx"
            cases = [excel_case(notes="补充备注")]
            build_valid_workbook("测试用例", cases, matrix_for_cases(*cases)).save(target)
            self.assertEqual(
                load_workbook(target)["测试用例"]["K3"].value,
                "[需求追溯] 来源：注册需求.md｜定位：2.1 手机号规则｜规则：手机号必须为11位数字\n补充备注",
            )

    def test_excel_supports_multiple_traceability_entries(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "multiple-traceability.xlsx"
            case = excel_case(traceability_refs=["TRACE-REG-002", "TRACE-REG-001"])
            matrix = multi_traceability_matrix()
            build_valid_workbook("测试用例", [case], matrix).save(target)
            self.assertEqual(validate_excel_file(target, matrix), [])
            self.assertEqual(
                load_workbook(target)["测试用例"]["K3"].value,
                "[需求追溯] 来源：注册需求.md｜定位：2.1 手机号规则｜规则：手机号必须为11位数字\n"
                "[需求追溯] 来源：注册需求.md｜定位：2.2 注册成功｜规则：注册成功后显示新注册信息\n"
                "无",
            )

    def test_excel_requirement_id_does_not_replace_traceability(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "missing-traceability.xlsx"
            cases = [excel_case(requirement_id="REQ-1")]
            matrix = matrix_for_cases(*cases)
            build_valid_workbook("测试用例", cases, matrix).save(target)
            workbook = load_workbook(target)
            workbook["测试用例"]["K3"] = "[需求追溯] REQ-1"
            workbook.save(target)
            errors = validate_excel_file(target, matrix)
            self.assertTrue(any("不得仅使用内部编号" in error for error in errors))

    def test_excel_requirement_id_is_optional_and_accepts_material_trace_id(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "blank-requirement-id.xlsx"
            case = excel_case(trace_id="US-1024")
            case.pop("requirement_id")
            matrix = matrix_for_cases(case)
            build_valid_workbook("测试用例", [case], matrix).save(target)
            self.assertEqual(validate_excel_file(target, matrix), [])
            self.assertEqual(load_workbook(target)["测试用例"]["A3"].value, "US-1024")

    def test_excel_traceability_requires_exact_label(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "legacy-label.xlsx"
            cases = [excel_case()]
            matrix = matrix_for_cases(*cases)
            build_valid_workbook("测试用例", cases, matrix).save(target)
            workbook = load_workbook(target)
            workbook["测试用例"]["K3"] = "追溯：来源：注册需求.md｜定位：2.1 手机号规则｜规则：手机号必须为11位数字"
            workbook.save(target)
            errors = validate_excel_file(target, matrix)
            self.assertTrue(any("[需求追溯] 来源：" in error for error in errors))

    def test_custom_sheet_validates_with_same_name(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "custom.xlsx"
            cases = [excel_case()]
            matrix = matrix_for_cases(*cases)
            build_valid_workbook("测试用例", cases, matrix, sheet_name="自定义").save(target)
            self.assertTrue(validate_excel_file(target, matrix))
            self.assertEqual(validate_excel_file(target, matrix, sheet_name="自定义"), [])

    def test_formula_like_input_remains_plain_text(self):
        markers = ["=1+1", "+1+1", "-1+1", "@SUM(1)"]
        case = excel_case(requirement_id=markers[0], id=markers[1], title=markers[2], type=markers[3])
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "text.xlsx"
            build_valid_workbook(
                "=标题",
                [case],
                matrix_for_cases(case),
                requirement_name="=需求",
                notes="@说明",
            ).save(target)
            workbook = load_workbook(target, data_only=False)
            sheet = workbook["测试用例"]
            self.assertEqual([sheet.cell(3, column).value for column in range(1, 5)], markers)
            self.assertTrue(all(sheet.cell(3, column).data_type == "s" for column in range(1, 12)))
            self.assertEqual(sheet["A1"].value, "=标题")
            self.assertEqual(sheet["A1"].data_type, "s")
            self.assertEqual(workbook["生成说明"]["B2"].value, "=需求")
            self.assertEqual(workbook["生成说明"]["B6"].value, "@说明")

    def test_excel_formula_cell_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "formula.xlsx"
            cases = [excel_case()]
            matrix = matrix_for_cases(*cases)
            build_valid_workbook("测试用例", cases, matrix).save(target)
            workbook = load_workbook(target)
            workbook["测试用例"]["H3"] = "=1+1"
            workbook.save(target)
            errors = validate_excel_file(target, matrix)
            self.assertTrue(any("测试数据”包含公式" in error for error in errors))

    def test_excel_freeze_panes_and_filter_required(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "layout.xlsx"
            cases = [excel_case()]
            matrix = matrix_for_cases(*cases)
            build_valid_workbook("测试用例", cases, matrix).save(target)
            workbook = load_workbook(target)
            sheet = workbook["测试用例"]
            sheet.freeze_panes = None
            sheet.auto_filter.ref = "A2:J3"
            workbook.save(target)
            errors = validate_excel_file(target, matrix)
            self.assertTrue(any("冻结窗格至 A3" in error for error in errors))
            self.assertTrue(any("筛选范围必须为 A2:K3" in error for error in errors))

    def test_excel_trailing_empty_styled_row_ignored_for_filter(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "trailing-row.xlsx"
            cases = [excel_case()]
            matrix = matrix_for_cases(*cases)
            build_valid_workbook("测试用例", cases, matrix).save(target)
            workbook = load_workbook(target)
            sheet = workbook["测试用例"]
            sheet.cell(row=4, column=1).number_format = "0.00"
            workbook.save(target)
            self.assertGreaterEqual(load_workbook(target)["测试用例"].max_row, 4)
            self.assertEqual(validate_excel_file(target, matrix), [])

    def test_excel_missing_case_id_reports_single_root_cause(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "missing-id.xlsx"
            cases = [excel_case()]
            matrix = matrix_for_cases(*cases)
            build_valid_workbook("测试用例", cases, matrix).save(target)
            workbook = load_workbook(target)
            workbook["测试用例"]["B3"] = None
            workbook.save(target)
            errors = validate_excel_file(target, matrix)
            self.assertTrue(any("缺少必填字段: '用例 ID'" in error for error in errors), errors)
            self.assertFalse(
                any("用例 ID “”未在需求追溯矩阵中关联" in error for error in errors),
                errors,
            )

    def test_excel_required_fields_duplicates_and_priority(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "invalid.xlsx"
            cases = [
                excel_case(priority="P5", preconditions=""),
                excel_case(title="提交合法注册"),
            ]
            matrix = matrix_for_cases(*cases)
            build_valid_workbook("测试用例", cases, matrix).save(target)
            workbook = load_workbook(target)
            workbook["测试用例"]["K3"] = "[需求追溯] REQ-001"
            workbook.save(target)
            errors = validate_excel_file(target, matrix)
            self.assertTrue(any("不得仅使用内部编号" in error for error in errors))
            self.assertTrue(any("缺少必填字段: '前置条件'" in error for error in errors))
            self.assertTrue(any("用例 ID 重复" in error for error in errors))
            self.assertTrue(any("用例标题重复" in error for error in errors))
            self.assertTrue(any("优先级无效" in error for error in errors))
            self.assertFalse(any("需求/追踪 ID" in error for error in errors))

    def test_excel_priority_is_case_sensitive(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "lowercase-priority.xlsx"
            cases = [excel_case(priority="p1")]
            matrix = matrix_for_cases(*cases)
            build_valid_workbook("测试用例", cases, matrix).save(target)
            errors = validate_excel_file(target, matrix)
            self.assertTrue(any("优先级无效" in error for error in errors), errors)

    def test_excel_type_design_method_and_stale_matrix_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "invalid-enum.xlsx"
            case = excel_case(type="异常测试", design_method="判定表")
            matrix = matrix_for_cases(case)
            matrix["TRACE-REG-001"]["case_ids"].append("TC-REG-STALE")
            build_valid_workbook("测试用例", [case], matrix).save(target)
            errors = validate_excel_file(target, matrix)
            self.assertTrue(any("用例类型无效" in error for error in errors))
            self.assertTrue(any("设计方法无效" in error for error in errors))
            self.assertTrue(any("TC-REG-STALE”未出现在工作表中" in error for error in errors))

    def test_excel_optional_post_cleanup_stays_in_notes(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "cleanup.xlsx"
            case = excel_case(notes="后置清理：删除测试账号并恢复配置。")
            matrix = matrix_for_cases(case)
            build_valid_workbook("测试用例", [case], matrix).save(target)
            self.assertEqual(validate_excel_file(target, matrix), [])
            notes = load_workbook(target)["测试用例"]["K3"].value
            self.assertTrue(notes.endswith("\n后置清理：删除测试账号并恢复配置。"))

    def test_excel_cli_round_trip(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "cases.json"
            target = Path(directory) / "cases.xlsx"
            matrix_path = Path(directory) / "matrix.json"
            cases = [excel_case()]
            matrix = matrix_for_cases(*cases)
            source.write_text(
                json.dumps(
                    {"meta": {"sheet_name": "自定义", **valid_meta()}, "cases": cases},
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            write_matrix(matrix_path, matrix)
            subprocess.run(
                [sys.executable, str(ROOT / "scripts" / "create_excel_cases.py"), "--output", str(target), "--input-json", str(source), "--traceability-matrix", str(matrix_path)],
                check=True,
                capture_output=True,
                env=subprocess_env(),
            )
            self.assertEqual(validate_excel_file(target, matrix, sheet_name="自定义"), [])

    def test_excel_generation_requires_force_to_overwrite(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "cases.json"
            target = Path(directory) / "cases.xlsx"
            matrix_path = Path(directory) / "matrix.json"
            cases = [excel_case()]
            write_matrix(matrix_path, matrix_for_cases(*cases))

            def write_source(title):
                source.write_text(
                    json.dumps(
                        {
                            "meta": {"title": title, "sheet_name": "测试用例", **valid_meta()},
                            "cases": cases,
                        },
                        ensure_ascii=False,
                    ),
                    encoding="utf-8",
                )

            command = [
                sys.executable,
                str(ROOT / "scripts" / "create_excel_cases.py"),
                "--output",
                str(target),
                "--input-json",
                str(source),
                "--traceability-matrix",
                str(matrix_path),
            ]

            write_source("旧标题")
            subprocess.run(
                command,
                check=True,
                capture_output=True,
                env=subprocess_env(),
            )
            original = target.read_bytes()

            write_source("新标题")
            blocked = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=subprocess_env(),
                check=False,
            )
            self.assertNotEqual(blocked.returncode, 0)
            self.assertIn("--force", blocked.stderr)
            self.assertEqual(target.read_bytes(), original)

            forced = subprocess.run(
                [*command, "--force"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=subprocess_env(),
                check=False,
            )
            self.assertEqual(forced.returncode, 0, forced.stderr)
            self.assertEqual(load_workbook(target)["测试用例"]["A1"].value, "新标题")

    def test_content_beyond_column_k_rejected(self):
        for value in ("额外内容", "=1+1"):
            with self.subTest(value=value):
                with tempfile.TemporaryDirectory() as directory:
                    target = Path(directory) / "extra-column.xlsx"
                    cases = [excel_case()]
                    matrix = matrix_for_cases(*cases)
                    build_valid_workbook("测试用例", cases, matrix).save(target)
                    workbook = load_workbook(target)
                    workbook["测试用例"]["L3"] = value
                    workbook.save(target)
                    errors = validate_excel_file(target, matrix)
                    self.assertTrue(any("只允许 A–K" in error for error in errors), errors)

    def test_blank_content_beyond_column_k_is_allowed(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "blank-extra-column.xlsx"
            cases = [excel_case()]
            matrix = matrix_for_cases(*cases)
            build_valid_workbook("测试用例", cases, matrix).save(target)
            workbook = load_workbook(target)
            workbook["测试用例"]["L3"] = "   "
            workbook.save(target)
            self.assertEqual(validate_excel_file(target, matrix), [])

    def test_excel_title_and_merge_required(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "title.xlsx"
            cases = [excel_case()]
            matrix = matrix_for_cases(*cases)

            build_valid_workbook("测试用例", cases, matrix).save(target)
            workbook = load_workbook(target)
            workbook["测试用例"]["A1"] = None
            workbook.save(target)
            errors = validate_excel_file(target, matrix)
            self.assertTrue(any("A1 标题不能为空" in error for error in errors), errors)

            build_valid_workbook("测试用例", cases, matrix).save(target)
            workbook = load_workbook(target)
            workbook["测试用例"].unmerge_cells("A1:K1")
            workbook.save(target)
            errors = validate_excel_file(target, matrix)
            self.assertTrue(any("A1:K1 合并" in error for error in errors), errors)

            build_valid_workbook("测试用例", cases, matrix).save(target)
            workbook = load_workbook(target)
            workbook["测试用例"].merge_cells("A10:B10")
            workbook.save(target)
            errors = validate_excel_file(target, matrix)
            self.assertTrue(any("恰好包含 A1:K1" in error for error in errors), errors)

    def test_excel_title_and_info_checked_before_missing_rows(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "empty-main.xlsx"
            cases = [excel_case()]
            matrix = matrix_for_cases(*cases)
            build_valid_workbook("测试用例", cases, matrix).save(target)
            workbook = load_workbook(target)
            sheet = workbook["测试用例"]
            sheet.delete_rows(2, sheet.max_row - 1)
            sheet["A1"] = None
            sheet.unmerge_cells("A1:K1")
            workbook.save(target)
            errors = validate_excel_file(target, matrix)
            self.assertTrue(any("A1 标题不能为空" in error for error in errors), errors)
            self.assertTrue(any("A1:K1 合并" in error for error in errors), errors)
            self.assertTrue(any("未发现有效用例数据行" in error for error in errors), errors)
            self.assertTrue(any("用例数量”为 1" in error for error in errors), errors)

    def test_excel_info_sheet_missing_and_labels(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "info-labels.xlsx"
            cases = [excel_case()]
            matrix = matrix_for_cases(*cases)

            build_valid_workbook("测试用例", cases, matrix).save(target)
            workbook = load_workbook(target)
            del workbook["生成说明"]
            workbook.save(target)
            errors = validate_excel_file(target, matrix)
            self.assertEqual(
                sum("生成说明" in error for error in errors),
                1,
                errors,
            )

            build_valid_workbook("测试用例", cases, matrix).save(target)
            workbook = load_workbook(target)
            info = workbook["生成说明"]
            info["A2"] = "覆盖范围"
            info["A6"] = "其他"
            workbook.save(target)
            errors = validate_excel_file(target, matrix)
            self.assertTrue(
                any("“需求名称”" in error and "缺少" in error for error in errors),
                errors,
            )
            self.assertTrue(
                any("“覆盖范围”" in error and "重复" in error for error in errors),
                errors,
            )
            self.assertTrue(
                any("“补充说明”" in error and "缺少" in error for error in errors),
                errors,
            )

    def test_excel_info_sheet_required_values_date_and_count(self):
        scenarios = (
            ({"B2": None}, "“需求名称”不能为空"),
            ({"B3": "2026-02-30"}, "上线日期"),
            ({"B3": "20260101"}, "上线日期"),
            ({"B4": 2}, "与实际用例行数 1 不一致"),
            ({"B4": "一"}, "必须为整数"),
        )
        for changes, expected in scenarios:
            with self.subTest(expected=expected):
                with tempfile.TemporaryDirectory() as directory:
                    target = Path(directory) / "info-values.xlsx"
                    cases = [excel_case()]
                    matrix = matrix_for_cases(*cases)
                    build_valid_workbook("测试用例", cases, matrix).save(target)
                    workbook = load_workbook(target)
                    info = workbook["生成说明"]
                    for cell, value in changes.items():
                        info[cell] = value
                    workbook.save(target)
                    errors = validate_excel_file(target, matrix)
                    self.assertTrue(any(expected in error for error in errors), errors)

    def test_generator_notes_array_is_joined(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "notes.xlsx"
            cases = [excel_case()]
            build_workbook(
                "测试用例",
                cases,
                matrix_for_cases(*cases),
                meta=valid_meta(notes=["说明1", "说明2"]),
            ).save(target)
            workbook = load_workbook(target)
            self.assertEqual(workbook["生成说明"]["B6"].value, "说明1\n说明2")

    def test_excel_info_sheet_formulas_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "info-formula.xlsx"
            cases = [excel_case()]
            matrix = matrix_for_cases(*cases)
            build_valid_workbook("测试用例", cases, matrix).save(target)
            workbook = load_workbook(target)
            workbook["生成说明"]["B2"] = "=A1"
            workbook.save(target)
            errors = validate_excel_file(target, matrix)
            self.assertTrue(any("生成说明第 2 行包含公式" in error for error in errors), errors)

    def test_excel_info_sheet_extra_columns_rejected(self):
        for value in ("额外内容", "=1+1"):
            with self.subTest(value=value):
                with tempfile.TemporaryDirectory() as directory:
                    target = Path(directory) / "info-extra-column.xlsx"
                    cases = [excel_case()]
                    matrix = matrix_for_cases(*cases)
                    build_valid_workbook("测试用例", cases, matrix).save(target)
                    workbook = load_workbook(target)
                    workbook["生成说明"]["C2"] = value
                    workbook.save(target)
                    errors = validate_excel_file(target, matrix)
                    self.assertTrue(any("只允许 A–B 两列" in error for error in errors), errors)

    def test_excel_title_and_header_formulas_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "header-formula.xlsx"
            cases = [excel_case()]
            matrix = matrix_for_cases(*cases)
            build_valid_workbook("测试用例", cases, matrix).save(target)
            workbook = load_workbook(target)
            sheet = workbook["测试用例"]
            sheet["A1"] = "=1+1"
            sheet["B2"] = "=1+1"
            workbook.save(target)
            errors = validate_excel_file(target, matrix)
            self.assertTrue(any("第 1 行标题包含公式" in error for error in errors), errors)
            self.assertTrue(any("第 2 行字段“用例 ID”包含公式" in error for error in errors), errors)

    def test_generator_meta_contract(self):
        cases = [excel_case()]
        matrix = matrix_for_cases(*cases)
        with self.assertRaisesRegex(ValueError, "requirement_name"):
            build_workbook("测试用例", cases, matrix)
        with self.assertRaisesRegex(ValueError, "requirement_name"):
            build_workbook("测试用例", cases, matrix, meta=valid_meta(requirement_name="  "))
        with self.assertRaisesRegex(ValueError, "release_date"):
            build_workbook("测试用例", cases, matrix, meta={"requirement_name": "注册需求"})
        with self.assertRaisesRegex(ValueError, "case_count"):
            build_workbook("测试用例", cases, matrix, meta=valid_meta(case_count=1))
        with self.assertRaisesRegex(ValueError, "case_count"):
            build_workbook("测试用例", cases, matrix, meta=valid_meta(case_count=None))
        with self.assertRaisesRegex(ValueError, "有效 YYYY-MM-DD"):
            build_workbook("测试用例", cases, matrix, meta=valid_meta(release_date="2026-02-30"))
        with self.assertRaisesRegex(ValueError, "有效 YYYY-MM-DD"):
            build_workbook("测试用例", cases, matrix, meta=valid_meta(release_date="20260101"))
        with self.assertRaisesRegex(ValueError, "标题不能为空"):
            build_workbook("   ", cases, matrix, meta=valid_meta())
        with self.assertRaisesRegex(ValueError, "cases.*非空数组"):
            build_workbook("测试用例", [], matrix, meta=valid_meta())
        with self.assertRaisesRegex(ValueError, "'meta' 必须是对象"):
            build_workbook("测试用例", cases, matrix, meta=[])

    def test_build_workbook_does_not_mutate_meta(self):
        cases = [excel_case()]
        meta = valid_meta()
        build_workbook("测试用例", cases, matrix_for_cases(*cases), meta=meta)
        self.assertEqual(meta, valid_meta())
        self.assertNotIn("case_count", meta)

    def test_excel_cli_accepts_utf8_bom_input(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "cases.json"
            target = Path(directory) / "cases.xlsx"
            matrix_path = Path(directory) / "matrix.json"
            cases = [excel_case()]
            matrix = matrix_for_cases(*cases)
            payload = json.dumps({"meta": valid_meta(), "cases": cases}, ensure_ascii=False)
            source.write_bytes(b"\xef\xbb\xbf" + payload.encode("utf-8"))
            write_matrix(matrix_path, matrix)
            result = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "create_excel_cases.py"),
                    "--output",
                    str(target),
                    "--input-json",
                    str(source),
                    "--traceability-matrix",
                    str(matrix_path),
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=subprocess_env(),
                check=False,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(validate_excel_file(target, matrix), [])

    def test_covered_entry_requires_excel_case_ids(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "missing-ids.xlsx"
            cases = [excel_case()]
            matrix = matrix_for_cases(*cases)
            build_valid_workbook("测试用例", cases, matrix).save(target)
            matrix["TRACE-REG-001"]["case_ids"] = []
            errors = validate_excel_file(target, matrix)
            self.assertTrue(
                any("缺少 Excel 用例 ID 关联" in error for error in errors),
                errors,
            )

    def test_excel_cli_default_title_from_meta(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "cases.json"
            target = Path(directory) / "cases.xlsx"
            matrix_path = Path(directory) / "matrix.json"
            cases = [excel_case()]
            write_matrix(matrix_path, matrix_for_cases(*cases))
            source.write_text(
                json.dumps({"meta": valid_meta(), "cases": cases}, ensure_ascii=False),
                encoding="utf-8",
            )
            result = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "create_excel_cases.py"),
                    "--output",
                    str(target),
                    "--input-json",
                    str(source),
                    "--traceability-matrix",
                    str(matrix_path),
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=subprocess_env(),
                check=False,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            workbook = load_workbook(target)
            self.assertEqual(
                workbook["测试用例"]["A1"].value,
                "测试用例 - 未指定日期 - [注册需求]",
            )

    def test_excel_cli_meta_error_has_no_traceback(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "cases.json"
            target = Path(directory) / "cases.xlsx"
            matrix_path = Path(directory) / "matrix.json"
            cases = [excel_case()]
            write_matrix(matrix_path, matrix_for_cases(*cases))
            source.write_text(
                json.dumps({"meta": {}, "cases": cases}, ensure_ascii=False),
                encoding="utf-8",
            )
            result = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "create_excel_cases.py"),
                    "--output",
                    str(target),
                    "--input-json",
                    str(source),
                    "--traceability-matrix",
                    str(matrix_path),
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=subprocess_env(),
                check=False,
            )
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("requirement_name", result.stderr)
            self.assertNotIn("Traceback", result.stderr)

    def test_excel_cli_missing_input_has_no_traceback(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "cases.xlsx"
            result = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "create_excel_cases.py"),
                    "--output",
                    str(target),
                    "--input-json",
                    str(Path(directory) / "missing-cases.json"),
                    "--traceability-matrix",
                    str(ROOT / "assets" / "example-traceability-matrix.json"),
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=subprocess_env(),
                check=False,
            )
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("无法读取输入 JSON", result.stderr)
            self.assertNotIn("Traceback", result.stderr)
            self.assertFalse(target.exists())

    def test_bundled_excel_template_metadata_is_filled(self):
        info = load_workbook(
            ROOT / "assets" / "excel-test-case-template.xlsx"
        )["生成说明"]
        self.assertEqual(info["B2"].value, "示例需求")
        self.assertEqual(info["B3"].value, "未指定日期")
        self.assertEqual(info["B4"].value, 1)

    def test_bundle_success_messages(self):
        with tempfile.TemporaryDirectory() as directory:
            markdown = Path(directory) / "cases.md"
            excel = Path(directory) / "cases.xlsx"
            matrix_path = Path(directory) / "matrix.json"
            matrix = matrix_for_cases(excel_case())
            markdown.write_text(VALID_MARKDOWN, encoding="utf-8")
            build_valid_workbook("测试用例", [excel_case()], matrix).save(excel)
            write_matrix(matrix_path, matrix)
            base = [
                sys.executable,
                str(ROOT / "scripts" / "validate_traceability_bundle.py"),
                "--traceability-matrix",
                str(matrix_path),
            ]
            scenarios = (
                (
                    [*base, "--markdown", str(markdown), "--excel", str(excel)],
                    "Markdown、Excel 格式及跨格式覆盖校验通过。",
                ),
                (
                    [*base, "--markdown", str(markdown)],
                    "Markdown 格式与矩阵关联校验通过。",
                ),
                (
                    [*base, "--excel", str(excel)],
                    "Excel 格式与矩阵关联校验通过。",
                ),
                (
                    [
                        *base,
                        "--markdown",
                        str(markdown),
                        "--excel",
                        str(excel),
                        "--coverage-only",
                    ],
                    "追溯覆盖核对通过（未执行 Markdown/Excel 格式校验）。",
                ),
            )
            for command, expected in scenarios:
                with self.subTest(expected=expected):
                    result = subprocess.run(
                        command,
                        capture_output=True,
                        text=True,
                        encoding="utf-8",
                        env=subprocess_env(),
                        check=False,
                    )
                    self.assertEqual(result.returncode, 0, result.stderr)
                    self.assertIn(expected, result.stdout)

    def test_excel_atomic_save_keeps_existing_file_on_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "cases.xlsx"
            target.write_bytes(b"original workbook")

            class FailingWorkbook:
                def save(self, path):
                    Path(path).write_bytes(b"partial workbook")
                    raise OSError("模拟保存失败")

            with self.assertRaisesRegex(OSError, "模拟保存失败"):
                save_workbook_atomic(FailingWorkbook(), target)

            self.assertEqual(target.read_bytes(), b"original workbook")
            self.assertEqual([path.name for path in Path(directory).iterdir()], ["cases.xlsx"])

    def test_bundled_excel_template(self):
        matrix, errors = load_traceability_matrix(
            ROOT / "assets" / "example-traceability-matrix.json"
        )
        self.assertEqual(errors, [])
        self.assertEqual(
            validate_excel_file(ROOT / "assets" / "excel-test-case-template.xlsx", matrix),
            [],
        )

    def test_bundle_validation_and_traceability_sync(self):
        with tempfile.TemporaryDirectory() as directory:
            markdown = Path(directory) / "cases.md"
            excel = Path(directory) / "cases.xlsx"
            matrix_path = Path(directory) / "matrix.json"
            matrix = multi_traceability_matrix()
            markdown.write_text(VALID_MARKDOWN, encoding="utf-8")
            case = excel_case(traceability_refs=["TRACE-REG-001", "TRACE-REG-002"])
            build_valid_workbook("测试用例", [case], matrix).save(excel)
            matrix["TRACE-REG-001"]["rule"] = "手机号必须为11位且不得包含空格"
            write_matrix(matrix_path, matrix)

            self.assertTrue(validate_bundle(matrix_path, markdown, excel))
            self.assertEqual(sync_atomically(markdown, excel, matrix), [])
            self.assertEqual(validate_bundle(matrix_path, markdown, excel), [])

    def test_sync_markdown_preserves_line_endings(self):
        matrix = {
            "TRACE-REG-001": {
                **TRACEABILITY_MATRIX["TRACE-REG-001"],
                "rule": "手机号必须为11位数字且不得包含空格",
            }
        }
        for newline, label in (("\n", "LF"), ("\r\n", "CRLF")):
            with self.subTest(label=label):
                with tempfile.TemporaryDirectory() as directory:
                    markdown = Path(directory) / "cases.md"
                    markdown.write_bytes(
                        VALID_MARKDOWN.replace("\n", newline).encode("utf-8")
                    )
                    self.assertEqual(sync_markdown(markdown, matrix), [])
                    data = markdown.read_bytes()
                    self.assertIn(
                        "手机号必须为11位数字且不得包含空格".encode("utf-8"),
                        data,
                    )
                    if newline == "\r\n":
                        self.assertEqual(data.count(b"\n"), data.count(b"\r\n"))
                    else:
                        self.assertNotIn(b"\r", data)

    def test_sync_atomically_leaves_both_files_unchanged_on_excel_error(self):
        with tempfile.TemporaryDirectory() as directory:
            markdown = Path(directory) / "cases.md"
            excel = Path(directory) / "cases.xlsx"
            matrix = multi_traceability_matrix()
            markdown.write_text(VALID_MARKDOWN, encoding="utf-8")
            build_valid_workbook(
                "测试用例",
                [excel_case(traceability_refs=["TRACE-REG-001", "TRACE-REG-002"])],
                matrix,
            ).save(excel)
            original_markdown = markdown.read_text(encoding="utf-8")
            original_excel = excel.read_bytes()
            matrix["TRACE-REG-001"]["rule"] = "手机号必须为11位且不得包含空格"

            errors = sync_atomically(markdown, excel, matrix, sheet_name="不存在的工作表")

            self.assertTrue(errors)
            self.assertEqual(markdown.read_text(encoding="utf-8"), original_markdown)
            self.assertEqual(excel.read_bytes(), original_excel)

    def test_sync_atomically_restores_committed_file_when_later_replace_fails(self):
        with tempfile.TemporaryDirectory() as directory:
            markdown = Path(directory) / "cases.md"
            excel = Path(directory) / "cases.xlsx"
            matrix = multi_traceability_matrix()
            markdown.write_text(VALID_MARKDOWN, encoding="utf-8")
            build_valid_workbook(
                "测试用例",
                [excel_case(traceability_refs=["TRACE-REG-001", "TRACE-REG-002"])],
                matrix,
            ).save(excel)
            original_markdown = markdown.read_bytes()
            original_excel = excel.read_bytes()
            matrix["TRACE-REG-001"]["rule"] = "手机号必须为11位且不得包含空格"

            real_replace = os.replace
            replace_calls: list[tuple[object, object]] = []

            def fail_second_replace(source, destination):
                replace_calls.append((source, destination))
                if len(replace_calls) == 2:
                    raise OSError("模拟第二次替换失败")
                return real_replace(source, destination)

            with mock.patch("sync_traceability.os.replace", side_effect=fail_second_replace):
                errors = sync_atomically(markdown, excel, matrix)

            self.assertTrue(any("模拟第二次替换失败" in error for error in errors))
            self.assertTrue(any("已恢复此前写入的文件" in error for error in errors))
            self.assertEqual(markdown.read_bytes(), original_markdown)
            self.assertEqual(excel.read_bytes(), original_excel)
            self.assertEqual(list(Path(directory).glob(".*")), [])

    def test_sync_atomically_reports_backup_when_rollback_fails(self):
        with tempfile.TemporaryDirectory() as directory:
            markdown = Path(directory) / "cases.md"
            excel = Path(directory) / "cases.xlsx"
            matrix = multi_traceability_matrix()
            markdown.write_text(VALID_MARKDOWN, encoding="utf-8")
            build_valid_workbook(
                "测试用例",
                [excel_case(traceability_refs=["TRACE-REG-001", "TRACE-REG-002"])],
                matrix,
            ).save(excel)
            original_markdown = markdown.read_bytes()
            original_excel = excel.read_bytes()
            matrix["TRACE-REG-001"]["rule"] = "手机号必须为11位且不得包含空格"

            real_replace = os.replace
            replace_calls: list[tuple[object, object]] = []

            def fail_after_first_replace(source, destination):
                replace_calls.append((source, destination))
                if len(replace_calls) >= 2:
                    raise OSError("模拟替换失败")
                return real_replace(source, destination)

            with mock.patch(
                "sync_traceability.os.replace",
                side_effect=fail_after_first_replace,
            ):
                errors = sync_atomically(markdown, excel, matrix)

            self.assertTrue(any("回滚 Excel 失败" in error for error in errors))
            self.assertTrue(any("请手工恢复" in error for error in errors))
            self.assertEqual(markdown.read_bytes(), original_markdown)
            leftovers = [path for path in Path(directory).iterdir() if path.name.startswith(".")]
            self.assertEqual(len(leftovers), 1)
            self.assertEqual(leftovers[0].read_bytes(), original_excel)
            self.assertIn(str(leftovers[0]), "\n".join(errors))

    def test_copy_to_temp_removes_partial_copy_on_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "cases.md"
            source.write_text(VALID_MARKDOWN, encoding="utf-8")

            with mock.patch(
                "sync_traceability.shutil.copy2",
                side_effect=OSError("模拟复制失败"),
            ):
                with self.assertRaisesRegex(OSError, "模拟复制失败"):
                    _copy_to_temp(source)

            self.assertEqual([path.name for path in Path(directory).iterdir()], ["cases.md"])

    def test_traceability_sync_preserves_non_traceability_quote(self):
        with tempfile.TemporaryDirectory() as directory:
            markdown = Path(directory) / "cases.md"
            note = "      > 执行说明：请在独立环境执行。\n"
            original = VALID_MARKDOWN.replace(
                "      > [需求追溯]", note + "      > [需求追溯]"
            )
            markdown.write_text(original, encoding="utf-8")
            errors = sync_markdown(markdown, TRACEABILITY_MATRIX)
            self.assertTrue(any("首个引用块不是需求追溯" in error for error in errors))
            self.assertEqual(markdown.read_text(encoding="utf-8"), original)

    def test_bundle_requires_each_covered_entry_in_both_formats(self):
        with tempfile.TemporaryDirectory() as directory:
            markdown = Path(directory) / "cases.md"
            excel = Path(directory) / "cases.xlsx"
            matrix_path = Path(directory) / "matrix.json"
            matrix = matrix_for_cases(excel_case())
            markdown.write_text(VALID_MARKDOWN, encoding="utf-8")
            build_valid_workbook("测试用例", [excel_case()], matrix).save(excel)
            matrix["TRACE-REG-001"]["case_titles"] = []
            write_matrix(matrix_path, matrix)
            errors = validate_bundle(matrix_path, markdown, excel)
            self.assertEqual(
                [error for error in errors if "缺少 Markdown 用例标题关联" in error],
                ["需求追溯矩阵中的已覆盖条目“TRACE-REG-001”缺少 Markdown 用例标题关联。"],
            )
            self.assertFalse(any("覆盖不一致" in error for error in errors), errors)

            matrix["TRACE-REG-001"]["case_titles"] = ["提交合法注册"]
            matrix["TRACE-REG-001"]["case_ids"] = []
            write_matrix(matrix_path, matrix)
            errors = validate_bundle(matrix_path, markdown, excel)
            self.assertEqual(
                [error for error in errors if "缺少 Excel 用例 ID 关联" in error],
                ["需求追溯矩阵中的已覆盖条目“TRACE-REG-001”缺少 Excel 用例 ID 关联。"],
            )
            self.assertFalse(any("覆盖不一致" in error for error in errors), errors)

    def test_bundle_coverage_only_skips_format_validation(self):
        with tempfile.TemporaryDirectory() as directory:
            markdown = Path(directory) / "cases.md"
            excel = Path(directory) / "cases.xlsx"
            matrix_path = Path(directory) / "matrix.json"
            matrix = matrix_for_cases(excel_case())
            invalid_markdown = VALID_MARKDOWN.replace(
                "      - 前置条件\n        - 用户未注册。\n", ""
            )
            markdown.write_text(invalid_markdown, encoding="utf-8")
            build_valid_workbook("测试用例", [excel_case()], matrix).save(excel)
            write_matrix(matrix_path, matrix)

            full_errors = validate_bundle(matrix_path, markdown, excel)
            self.assertTrue(any("缺少必填字段“前置条件”" in error for error in full_errors))
            self.assertEqual(
                validate_bundle(matrix_path, markdown, excel, coverage_only=True),
                [],
            )

            result = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "validate_traceability_bundle.py"),
                    "--traceability-matrix",
                    str(matrix_path),
                    "--markdown",
                    str(markdown),
                    "--excel",
                    str(excel),
                    "--coverage-only",
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=subprocess_env(),
                check=False,
            )
            self.assertEqual(result.returncode, 0, result.stderr)

            missing_format = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "validate_traceability_bundle.py"),
                    "--traceability-matrix",
                    str(matrix_path),
                    "--markdown",
                    str(markdown),
                    "--coverage-only",
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=subprocess_env(),
                check=False,
            )
            self.assertNotEqual(missing_format.returncode, 0)
            self.assertIn("必须同时提供", missing_format.stderr)

            missing_paths = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "validate_traceability_bundle.py"),
                    "--traceability-matrix",
                    str(matrix_path),
                    "--markdown",
                    str(Path(directory) / "missing.md"),
                    "--excel",
                    str(Path(directory) / "missing.xlsx"),
                    "--coverage-only",
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=subprocess_env(),
                check=False,
            )
            self.assertNotEqual(missing_paths.returncode, 0)
            self.assertIn("Markdown 文件不存在", missing_paths.stderr)
            self.assertIn("Excel 文件不存在", missing_paths.stderr)

    def test_bundle_coverage_only_requires_both_format_links(self):
        with tempfile.TemporaryDirectory() as directory:
            markdown = Path(directory) / "cases.md"
            excel = Path(directory) / "cases.xlsx"
            matrix_path = Path(directory) / "matrix.json"
            matrix = matrix_for_cases(excel_case())
            matrix["TRACE-REG-001"]["case_titles"] = []
            markdown.write_text(VALID_MARKDOWN, encoding="utf-8")
            build_valid_workbook("测试用例", [excel_case()], matrix).save(excel)
            write_matrix(matrix_path, matrix)

            errors = validate_bundle(
                matrix_path,
                markdown,
                excel,
                coverage_only=True,
            )
            self.assertTrue(any("缺少 Markdown 用例标题关联" in error for error in errors))
            self.assertTrue(
                any("仅覆盖核对必须同时提供" in error for error in validate_bundle(
                    matrix_path,
                    markdown_path=markdown,
                    coverage_only=True,
                ))
            )

            matrix["TRACE-REG-001"]["status"] = "待覆盖"
            matrix["TRACE-REG-001"]["case_ids"] = []
            write_matrix(matrix_path, matrix)
            draft_errors = validate_bundle(
                matrix_path,
                markdown,
                excel,
                coverage_only=True,
            )
            self.assertTrue(any("存在待覆盖条目" in error for error in draft_errors))

    def test_markdown_template_uses_matrix_as_the_only_pending_source(self):
        template = (ROOT / "assets" / "md-test-case-template.md").read_text(encoding="utf-8")
        self.assertNotIn("需求存疑与待确认清单", template)
        self.assertNotIn("[待确认-", template)
        self.assertNotIn("正式需求编号", template)
        self.assertIn("{材料文件}", template)


class CaseBaselineTests(unittest.TestCase):
    def test_validate_accepts_minimal_ir(self):
        ir = valid_ir()
        self.assertEqual(validate_baseline(ir), [])
        self.assertEqual(validate_baseline(ir, require_design_method=True), [])

    def test_validate_rejects_meta_violations(self):
        for mutate, expected in (
            (lambda ir: ir["meta"].update(case_count=1), "case_count"),
            (lambda ir: ir["meta"].update(requirement_name=" "), "requirement_name"),
            (lambda ir: ir["meta"].update(release_date="2026-02-30"), "release_date"),
            (lambda ir: ir["meta"].update(sheet_name="需求追溯"), "保留名"),
        ):
            with self.subTest(expected=expected):
                ir = valid_ir()
                mutate(ir)
                self.assertTrue(
                    any(expected in error for error in validate_baseline(ir)),
                    validate_baseline(ir),
                )

    def test_validate_rejects_title_pollution_and_duplicates(self):
        for title, expected in (
            ("正向-提交注册", "不得以"),
            ("提交注册-P1", "不得以"),
        ):
            with self.subTest(title=title):
                ir = valid_ir()
                ir["cases"][0]["title"] = title
                self.assertTrue(
                    any(expected in error for error in validate_baseline(ir)),
                    validate_baseline(ir),
                )

        ir = valid_ir()
        duplicate = deepcopy(ir["cases"][0])
        duplicate["title"] = "另一个标题"
        ir["cases"].append(duplicate)
        ir["traceability"][0]["case_ids"] = ["TC-ZC-001"]
        self.assertTrue(any("用例 ID 重复" in error for error in validate_baseline(ir)))

    def test_validate_requires_valid_design_method_for_excel(self):
        ir = valid_ir()
        ir["cases"][0]["design_method"] = ""
        self.assertTrue(
            any("design_method" in error for error in validate_baseline(ir, require_design_method=True))
        )
        self.assertFalse(any("design_method" in error for error in validate_baseline(ir)))

        ir = valid_ir()
        ir["cases"][0]["design_method"] = "判定表"
        self.assertTrue(
            any("design_method" in error for error in validate_baseline(ir, require_design_method=True))
        )

    def test_validate_rejects_pending_reference_and_broken_links(self):
        ir = valid_ir()
        ir["traceability"].append(
            {
                "id": "TRACE-REG-Q01",
                "source": "注册需求.md",
                "location": "1.2 超时",
                "status": "待确认",
                "issue": "超时行为未定义",
                "assumption": "暂按失败处理",
                "question": "超时是否重试？",
            }
        )
        ir["cases"][0]["traceability_refs"] = ["TRACE-REG-Q01"]
        errors = validate_baseline(ir)
        self.assertTrue(any("不得引用待确认条目" in error for error in errors), errors)

        ir = valid_ir()
        ir["traceability"][0]["case_ids"] = ["TC-ZC-999"]
        self.assertTrue(any("不一致" in error for error in validate_baseline(ir)))

    def test_assign_ids_preserves_gaps_and_uses_max_plus_one(self):
        cases = [{"id": "TC-ZC-001"}, {"id": "TC-ZC-003"}, {"id": "TC-NEW"}]
        assigned = assign_missing_ids(cases)
        self.assertEqual(assigned, ["TC-ZC-004"])
        self.assertNotIn("TC-ZC-002", [case["id"] for case in cases])

    def test_renumber_continuous_for_single_prefix_and_buckets_for_multi(self):
        cases = [{"id": "TC-ZC-005"}, {"id": "TC-ZC-009"}]
        traceability = [{"id": "TRACE-1", "case_ids": ["TC-ZC-005", "TC-ZC-009"]}]
        mapping = renumber_cases(cases, traceability)
        self.assertEqual([case["id"] for case in cases], ["TC-ZC-001", "TC-ZC-002"])
        self.assertEqual(traceability[0]["case_ids"], ["TC-ZC-001", "TC-ZC-002"])
        self.assertEqual(mapping["TC-ZC-009"], "TC-ZC-002")

        cases = [
            {"id": "TC-A-005", "module": "M", "type": "功能测试", "group": ""},
            {"id": "TC-B-009", "module": "M", "type": "功能测试", "group": ""},
            {"id": "TC-A-010", "module": "M", "type": "功能测试", "group": ""},
        ]
        renumber_cases(cases, [])
        self.assertEqual(
            [case["id"] for case in cases], ["TC-A-001", "TC-B-001", "TC-A-002"]
        )

        cases = [{"id": "TC-A-005", "module": "M", "type": "功能测试", "group": ""}]
        renumber_cases(cases, [], id_prefix="TC-Z")
        self.assertEqual(cases[0]["id"], "TC-Z-001")

    def test_canonical_tree_order_groups_by_first_appearance(self):
        cases = [
            {"id": "1", "module": "B", "type": "功能测试", "group": ""},
            {"id": "2", "module": "A", "type": "功能测试", "group": ""},
            {"id": "3", "module": "B", "type": "功能测试", "group": ""},
        ]
        self.assertEqual([case["id"] for case in canonical_tree_order(cases)], ["1", "3", "2"])

    def test_canonical_tree_order_nests_groups_like_the_rendered_tree(self):
        cases = [
            {"id": "C1", "module": "M1", "type": "功能测试", "group": "G1"},
            {"id": "C2", "module": "M2", "type": "功能测试", "group": "G1"},
            {"id": "C3", "module": "M1", "type": "功能测试", "group": "G2"},
        ]
        self.assertEqual(
            [case["id"] for case in canonical_tree_order(cases)], ["C1", "C3", "C2"]
        )

    def test_markdown_and_excel_render_the_same_case_order(self):
        ir = valid_ir()
        template = ir["cases"][0]
        ir["cases"] = []
        for index, (module, group) in enumerate(
            (("M1", "G1"), ("M2", "G1"), ("M1", "G2")), start=1
        ):
            case = deepcopy(template)
            case.update(
                id=f"TC-ZC-00{index}",
                module=module,
                group=group,
                title=f"用例{index}",
            )
            ir["cases"].append(case)
        ir["traceability"][0]["case_ids"] = ["TC-ZC-001", "TC-ZC-003", "TC-ZC-002"]

        text = render_markdown(normalize_baseline(ir))
        markdown_order = []
        for line in text.splitlines():
            stripped = line.strip()
            if stripped.startswith("> [TC-"):
                markdown_order.append(stripped[3 : stripped.index("]")])
        sheet = build_workbook_from_ir(ir)["测试用例"]
        excel_order = [
            sheet.cell(row=row, column=2).value for row in range(3, 6)
        ]
        self.assertEqual(markdown_order, ["TC-ZC-001", "TC-ZC-003", "TC-ZC-002"])
        self.assertEqual(excel_order, markdown_order)

    def test_renumber_keeps_case_ids_in_case_order(self):
        cases = [
            {"id": "TC-A-001", "module": "B", "type": "功能测试", "group": ""},
            {"id": "TC-A-002", "module": "A", "type": "功能测试", "group": ""},
            {"id": "TC-A-003", "module": "B", "type": "功能测试", "group": ""},
        ]
        entries = [
            {
                "id": "TRACE-1",
                "source": "需求.md",
                "location": "1.1",
                "rule": "规则",
                "status": "已覆盖",
                "case_ids": ["TC-A-001", "TC-A-002", "TC-A-003"],
            }
        ]
        renumber_cases(cases, entries)
        self.assertEqual(entries[0]["case_ids"], [case["id"] for case in cases])

    def test_merge_with_excel_only_and_unreadable_source(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            excel = directory_path / "cases.xlsx"
            save_workbook_atomic(build_workbook_from_ir(valid_ir()), excel)

            merged, errors, _ = merge_artifacts_to_ir(excel_path=excel)
            self.assertEqual(errors, [])
            self.assertEqual(merged["cases"][0]["id"], "TC-ZC-001")

            merged, errors, _ = merge_artifacts_to_ir(
                md_path=directory_path / "missing.md", excel_path=excel
            )
            self.assertEqual(merged, {})
            self.assertTrue(any("无法读取 Markdown" in error for error in errors), errors)

    def test_merge_reports_missing_fields_after_combining_sources(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            markdown = directory_path / "cases.md"
            excel = directory_path / "cases.xlsx"
            ir = normalize_baseline(valid_ir())
            markdown.write_text(render_markdown(ir), encoding="utf-8")
            save_workbook_atomic(build_workbook_from_ir(valid_ir()), excel)

            _, errors, warnings = merge_artifacts_to_ir(
                md_path=markdown, excel_path=excel
            )
        self.assertEqual(errors, [])
        self.assertNotIn(
            "Markdown 不包含设计方法，导出 IR 后需从 Excel 合并或手工补充 design_method。",
            warnings,
        )
        self.assertFalse(
            any("缺少 design_method" in warning for warning in warnings), warnings
        )
        self.assertFalse(
            any("不包含正向/反向" in warning for warning in warnings), warnings
        )

    def test_merge_duplicate_entries_only_collapses_the_same_id(self):
        entries = [
            {
                "id": "TRACE-1",
                "source": "s",
                "location": "l",
                "rule": "r",
                "status": "已覆盖",
                "case_ids": ["TC-1"],
            },
            {
                "id": "TRACE-2",
                "source": "s",
                "location": "l",
                "rule": "r",
                "status": "已覆盖",
                "case_ids": ["TC-2"],
            },
            {
                "id": "TRACE-Q1",
                "source": "s",
                "location": "q",
                "status": "待确认",
                "issue": "i",
                "assumption": "a",
                "question": "q",
            },
        ]
        warnings: list[str] = []
        merged = _merge_duplicate_entries(entries, warnings)
        self.assertEqual([entry["id"] for entry in merged], ["TRACE-1", "TRACE-2", "TRACE-Q1"])
        self.assertNotIn("case_ids", merged[2])
        self.assertEqual(warnings, [])

    def test_markdown_round_trip_preserves_core_fields(self):
        ir = normalize_baseline(valid_ir())
        text = render_markdown(ir)
        self.assertIn("> [TC-ZC-001] 追溯: TRACE-REG-001", text)
        self.assertNotIn("### 待确认项", text)
        self.assertLess(text.index("- 前置操作"), text.index("- 测试数据"))
        self.assertLess(text.index("- 测试数据"), text.index("- 测试步骤"))
        self.assertLess(text.index("- 测试步骤"), text.index("- 预期结果"))
        self.assertLess(text.index("- 预期结果"), text.index("- 后置操作"))

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cases.md"
            path.write_text(text, encoding="utf-8")
            parsed, errors, _ = parse_markdown_to_ir(path)
        self.assertEqual(errors, [])
        original = ir["cases"][0]
        restored = parsed["cases"][0]
        for field in (
            "id",
            "module",
            "type",
            "polarity",
            "title",
            "priority",
            "pre_operations",
            "test_data",
            "steps",
            "expected_results",
            "post_operations",
            "traceability_refs",
        ):
            self.assertEqual(original[field], restored[field], field)
        self.assertEqual(parsed["traceability"][0]["case_ids"], ["TC-ZC-001"])

    def test_markdown_render_includes_pending_section(self):
        ir = valid_ir()
        ir["traceability"].append(
            {
                "id": "TRACE-REG-Q01",
                "source": "注册需求.md",
                "location": "1.2 超时",
                "status": "待确认",
                "issue": "超时行为未定义",
                "assumption": "暂按失败处理",
                "question": "超时是否重试？",
            }
        )
        text = render_markdown(normalize_baseline(ir))
        self.assertIn("### 待确认项", text)
        self.assertIn("超时是否重试？", text)

    def test_markdown_parser_merges_continuation_and_warns(self):
        text = (
            "# 测试用例 - 2026-11-24 - [注册]\n\n"
            "> 用例数量: 用例 1 条\n\n"
            "- [ ] 注册模块\n\n"
            "  - [ ] 功能测试\n\n"
            "    - [ ] 正向-[注册模块]_[注册]_[成功]-P0\n\n"
            "      > [TC-ZC-001] 追溯: TRACE-1\n\n"
            "      - 前置操作\n\n"
            "        - 用户未注册。\n\n"
            "      - 测试步骤\n\n"
            "        - 提交注册。\n"
            "          由脚本自动提交。\n\n"
            "      - 预期结果\n\n"
            "        - 注册成功。\n\n"
            "## 需求追溯\n\n"
            "### 需求追溯表\n\n"
            "| 追溯 ID | 来源 | 定位 | 规则 | 关联用例 |\n"
            "| :--- | :--- | :--- | :--- | :--- |\n"
            "| TRACE-1 | 需求.md | 1.1 | 规则 | TC-ZC-001 |\n"
        )
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cases.md"
            path.write_text(text, encoding="utf-8")
            parsed, errors, warnings = parse_markdown_to_ir(path)
        self.assertEqual(errors, [])
        self.assertEqual(
            parsed["cases"][0]["steps"], ["提交注册。；由脚本自动提交。"]
        )
        self.assertTrue(any("合并" in warning for warning in warnings), warnings)

    def test_markdown_parser_rejects_legacy_fields_and_pollution(self):
        base = render_markdown(normalize_baseline(valid_ir()))
        cases = (
            (base.replace("- 测试数据", "- 备注"), "v2 已删除字段"),
            (base.replace("## 需求追溯", "## 其他小节"), "缺少唯一"),
            (base.replace("正向-[注册模块]", "正向-TC-ZC-099-[注册模块]"), "不得包含用例 ID"),
            (base.replace("> [TC-ZC-001] 追溯", "> 追溯"), "缺少用例 ID"),
        )
        for text, expected in cases:
            with self.subTest(expected=expected):
                with tempfile.TemporaryDirectory() as directory:
                    path = Path(directory) / "cases.md"
                    path.write_text(text, encoding="utf-8")
                    _, errors, _ = parse_markdown_to_ir(path)
                self.assertTrue(any(expected in error for error in errors), errors)

    def test_excel_v2_round_trip(self):
        workbook = build_workbook_from_ir(valid_ir())
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cases.xlsx"
            save_workbook_atomic(workbook, path)
            parsed, errors, _ = parse_excel_to_ir(path)
        self.assertEqual(errors, [])
        case = parsed["cases"][0]
        self.assertEqual(case["id"], "TC-ZC-001")
        self.assertEqual(case["design_method"], "等价类划分")
        self.assertEqual(case["steps"], ["提交合法手机号。"])
        self.assertEqual(case["traceability_refs"], ["TRACE-REG-001"])
        self.assertEqual(case["polarity"], "")

    def test_excel_v1_requires_and_uses_legacy_matrix(self):
        case = excel_case()
        workbook = build_workbook(
            "测试用例", [case], TRACEABILITY_MATRIX, meta=valid_meta()
        )
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "legacy.xlsx"
            matrix_path = Path(directory) / "matrix.json"
            save_workbook_atomic(workbook, target)
            write_matrix(matrix_path, TRACEABILITY_MATRIX)
            _, errors, _ = parse_excel_to_ir(target)
            self.assertTrue(any("--legacy-matrix" in error for error in errors), errors)
            parsed, errors, _ = parse_excel_to_ir(
                target, legacy_matrix_path=matrix_path
            )
        self.assertEqual(errors, [])
        self.assertEqual(parsed["cases"][0]["traceability_refs"], ["TRACE-REG-001"])

    def test_validate_reads_raw_fields_before_normalizing(self):
        ir = valid_ir()
        ir["meta"]["case_count"] = 99
        self.assertTrue(any("case_count" in error for error in validate_baseline(ir)))

        ir = valid_ir()
        ir["cases"][0]["traceability_refs"] = ["TRACE-REG-001", "TRACE-REG-001"]
        self.assertTrue(any("重复项" in error for error in validate_baseline(ir)))

        ir = valid_ir()
        ir["meta"]["notes"] = "[待确认] 备注"
        self.assertTrue(
            any("meta 不得保留" in error for error in validate_baseline(ir))
        )

        ir = valid_ir()
        ir["cases"][0]["requirement_id"] = "REQ-1"
        ir["cases"][0]["notes"] = "旧备注"
        errors = validate_baseline(ir)
        self.assertTrue(any("requirement_id" in error for error in errors), errors)
        self.assertTrue(any("notes" in error for error in errors), errors)

    def test_draft_validation_skips_unassigned_links(self):
        ir = valid_ir()
        ir["traceability"][0]["status"] = "待覆盖"
        ir["traceability"][0]["case_ids"] = []
        self.assertEqual(
            validate_baseline(ir, require_finalized_traceability=False), []
        )
        self.assertTrue(validate_baseline(ir))

        ir["cases"][0]["id"] = "TC-NEW"
        self.assertEqual(
            validate_baseline(
                ir, require_finalized_traceability=False, allow_unassigned_ids=True
            ),
            [],
        )

    def test_markdown_parser_reports_exact_line_and_merges_sub_items(self):
        text = (
            "# 测试用例 - 2026-11-24 - [注册]\n\n"
            "> 用例数量: 用例 1 条\n\n"
            "- [ ] 注册模块\n\n"
            "  - [ ] 功能测试\n\n"
            "    - [ ] 正向-[注册模块]_[成功]-P0\n\n"
            "      > [TC-ZC-001] 追溯: TRACE-1\n\n"
            "      - 前置操作\n\n"
            "        - 用户未注册。\n\n"
            "      - 测试步骤\n\n"
            "        - 提交注册。\n"
            "          - 补充子项说明。\n\n"
            "      - 预期结果\n\n"
            "        - 注册成功。\n\n"
            "## 需求追溯\n\n"
            "### 需求追溯表\n\n"
            "| 追溯 ID | 来源 | 定位 | 规则 | 关联用例 |\n"
            "| :--- | :--- | :--- | :--- | :--- |\n"
            "| TRACE-1 | 需求.md | 1.1 | 规则 | TC-ZC-001 |\n"
        )
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cases.md"
            path.write_text(text, encoding="utf-8")
            parsed, errors, warnings = parse_markdown_to_ir(path)
        self.assertEqual(errors, [])
        self.assertEqual(parsed["cases"][0]["steps"], ["提交注册。；补充子项说明。"])
        self.assertTrue(any("第 20 行" in warning for warning in warnings), warnings)

    def test_duplicate_field_does_not_pollute_previous_field(self):
        base = render_markdown(normalize_baseline(valid_ir()))
        marker = "      - 预期结果"
        text = base.replace(
            marker, "      - 测试步骤\n\n        - 被吞掉的步骤。\n\n" + marker, 1
        )
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cases.md"
            path.write_text(text, encoding="utf-8")
            parsed, errors, _ = parse_markdown_to_ir(path)
        self.assertTrue(any("重复出现" in error for error in errors), errors)
        self.assertEqual(parsed["cases"][0]["steps"], ["提交合法手机号。"])

    def test_traceability_section_strictness(self):
        base = render_markdown(normalize_baseline(valid_ir()))
        pending_header = (
            "| 追溯 ID | 来源 | 定位 | 未明确项 | 当前测试假设 | 确认问题 |\n"
        )
        empty_separator = (
            "\n### 待确认项\n\n"
            + pending_header
            + "| | | | | | |\n"
        )
        empty_table = (
            "\n### 待确认项\n\n"
            + pending_header
            + "| :--- | :--- | :--- | :--- | :--- | :--- |\n"
        )
        unknown_heading = "\n### 其他小节\n\n| a | b |\n| - | - |\n"
        for suffix, expected in (
            (empty_separator, "缺少标准分隔行"),
            (empty_table, "表格不得为空"),
            (unknown_heading, "缺少 ### 需求追溯表"),
        ):
            with self.subTest(expected=expected):
                with tempfile.TemporaryDirectory() as directory:
                    path = Path(directory) / "cases.md"
                    path.write_text(base + suffix, encoding="utf-8")
                    _, errors, _ = parse_markdown_to_ir(path)
                self.assertTrue(
                    any(expected in error for error in errors), errors
                )

    def test_traceability_section_requires_its_tables(self):
        base = render_markdown(normalize_baseline(valid_ir()))
        cases = (
            (base.split("## 需求追溯")[0] + "## 需求追溯\n", "缺少 ### 需求追溯表"),
            (base + "\n### 待确认项\n", "标题后缺少表格"),
        )
        for text, expected in cases:
            with self.subTest(expected=expected):
                with tempfile.TemporaryDirectory() as directory:
                    path = Path(directory) / "cases.md"
                    path.write_text(text, encoding="utf-8")
                    _, errors, _ = parse_markdown_to_ir(path)
                self.assertTrue(any(expected in error for error in errors), errors)

    def test_render_cli_validates_the_raw_ir(self):
        import render_markdown_cases

        ir = valid_ir()
        ir["meta"]["case_count"] = 99
        with tempfile.TemporaryDirectory() as directory:
            baseline = Path(directory) / "ir.json"
            baseline.write_text(json.dumps(ir, ensure_ascii=False), encoding="utf-8")
            code = render_markdown_cases.main(
                ["--baseline", str(baseline), "--output", str(Path(directory) / "c.md")]
            )
        self.assertEqual(code, 1)

    def test_clis_report_locked_files_instead_of_raising(self):
        import case_baseline
        import create_excel_cases
        import render_markdown_cases

        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            baseline = directory_path / "ir.json"
            baseline.write_text(
                json.dumps(valid_ir(), ensure_ascii=False), encoding="utf-8"
            )
            locked = FileLockedError("文件被占用，无法写入。")

            with mock.patch.object(
                render_markdown_cases, "write_text_atomic", side_effect=locked
            ):
                self.assertEqual(
                    render_markdown_cases.main(
                        [
                            "--baseline",
                            str(baseline),
                            "--output",
                            str(directory_path / "cases.md"),
                        ]
                    ),
                    1,
                )

            with mock.patch.object(
                case_baseline, "write_text_atomic", side_effect=locked
            ):
                self.assertEqual(
                    case_baseline.main(["assign-ids", "--baseline", str(baseline)]), 1
                )

            argv = [
                "create_excel_cases.py",
                "--baseline",
                str(baseline),
                "--output",
                str(directory_path / "cases.xlsx"),
            ]
            with mock.patch.object(
                create_excel_cases, "save_workbook_atomic", side_effect=locked
            ), mock.patch.object(sys, "argv", argv):
                with self.assertRaises(SystemExit) as context:
                    create_excel_cases.main()
        self.assertIn("文件被占用", str(context.exception))


class ExcelV2RenderTests(unittest.TestCase):
    def test_workbook_layout_and_row_heights(self):
        workbook = build_workbook_from_ir(valid_ir())
        self.assertEqual(workbook.sheetnames, ["测试用例", "生成说明", "需求追溯"])
        sheet = workbook["测试用例"]
        self.assertEqual(
            [sheet.cell(row=2, column=column).value for column in range(1, 12)],
            [
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
            ],
        )
        self.assertEqual(sheet["A3"].value, "TRACE-REG-001")
        self.assertEqual(sheet["B3"].value, "TC-ZC-001")
        self.assertEqual(sheet["G3"].value, "用户未注册。")
        self.assertEqual(sheet.freeze_panes, "A3")
        self.assertEqual(sheet.auto_filter.ref, "A2:K3")
        self.assertEqual(sheet.column_dimensions["A"].width, 22)
        self.assertEqual(sheet["A3"].data_type, "s")
        height = sheet.row_dimensions[3].height
        self.assertGreaterEqual(height, 24)
        self.assertLessEqual(height, 409)

    def test_formula_like_text_stays_plain_and_real_formulas_are_rejected(self):
        ir = valid_ir()
        ir["cases"][0]["test_data"] = ["=1+1"]
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cases.xlsx"
            save_workbook_atomic(build_workbook_from_ir(ir), path)
            parsed, errors, _ = parse_excel_to_ir(path)
            self.assertEqual(errors, [])
            self.assertEqual(parsed["cases"][0]["test_data"], ["=1+1"])

            workbook = build_workbook_from_ir(valid_ir())
            workbook["测试用例"]["G3"] = "=1+1"
            save_workbook_atomic(workbook, path)
            _, errors, _ = parse_excel_to_ir(path)
        self.assertTrue(any("包含公式" in error for error in errors), errors)

    def test_info_sheet_uses_shared_row_height_and_strips_sheet_name(self):
        ir = valid_ir()
        ir["meta"]["coverage"] = "a\nb\nc"
        workbook = build_workbook_from_ir(ir)
        self.assertEqual(workbook["生成说明"].row_dimensions[5].height, 56)
        self.assertEqual(
            build_workbook_from_ir(valid_ir(), sheet_name=" 测试用例 ").sheetnames[0],
            "测试用例",
        )
        self.assertEqual(
            build_workbook_from_ir(valid_ir(), sheet_name="   ").sheetnames[0],
            "测试用例",
        )

    def test_pending_sheet_is_conditional(self):
        ir = valid_ir()
        workbook = build_workbook_from_ir(ir)
        self.assertNotIn("待确认项", workbook.sheetnames)

        ir["traceability"].append(
            {
                "id": "TRACE-REG-Q01",
                "source": "注册需求.md",
                "location": "1.2 超时",
                "status": "待确认",
                "issue": "超时行为未定义",
                "assumption": "暂按失败处理",
                "question": "超时是否重试？",
            }
        )
        workbook = build_workbook_from_ir(ir)
        self.assertEqual(
            workbook.sheetnames, ["测试用例", "生成说明", "需求追溯", "待确认项"]
        )
        pending = workbook["待确认项"]
        self.assertEqual(pending["A3"].value, "TRACE-REG-Q01")
        self.assertEqual(pending["F3"].value, "超时是否重试？")

    def test_reserved_sheet_name_and_missing_design_method_rejected(self):
        ir = valid_ir()
        ir["meta"]["sheet_name"] = "需求追溯"
        with self.assertRaisesRegex(ValueError, "保留名"):
            build_workbook_from_ir(ir)

        ir = valid_ir()
        ir["cases"][0]["design_method"] = ""
        with self.assertRaisesRegex(ValueError, "design_method"):
            build_workbook_from_ir(ir)

    def test_cli_renders_both_formats_from_ir(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            baseline = directory_path / "ir.json"
            markdown = directory_path / "cases.md"
            excel = directory_path / "cases.xlsx"
            baseline.write_text(
                json.dumps(valid_ir(), ensure_ascii=False), encoding="utf-8"
            )
            env = subprocess_env()
            md_result = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "render_markdown_cases.py"),
                    "--baseline",
                    str(baseline),
                    "--output",
                    str(markdown),
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=env,
            )
            self.assertEqual(md_result.returncode, 0, md_result.stderr)
            xlsx_result = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "create_excel_cases.py"),
                    "--baseline",
                    str(baseline),
                    "--output",
                    str(excel),
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=env,
            )
            self.assertEqual(xlsx_result.returncode, 0, xlsx_result.stderr)
            self.assertIn("TC-ZC-001", markdown.read_text(encoding="utf-8"))
            self.assertEqual(
                load_workbook(excel).sheetnames,
                ["测试用例", "生成说明", "需求追溯"],
            )


class ValidationUtilsTests(unittest.TestCase):
    def test_accepts_iso_date_and_placeholder(self):
        self.assertTrue(is_valid_date_or_placeholder("2026-11-19"))
        self.assertTrue(is_valid_date_or_placeholder(" 2026-11-19 "))
        self.assertTrue(is_valid_date_or_placeholder("未指定日期"))

    def test_rejects_compact_invalid_and_empty(self):
        for value in ("20261119", "2026-02-30", "2026-1-1", "", None, "2026/11/19"):
            with self.subTest(value=value):
                self.assertFalse(is_valid_date_or_placeholder(value))

    def test_configure_utf8_stdio_fixes_redirected_streams(self):
        code = (
            "import sys\n"
            f"sys.path.insert(0, {str(ROOT / 'scripts')!r})\n"
            "from validation_utils import configure_utf8_stdio\n"
            "configure_utf8_stdio()\n"
            "print('中文标准输出')\n"
            "print('中文标准错误', file=sys.stderr)\n"
        )
        env = os.environ.copy()
        for key in ("PYTHONIOENCODING", "PYTHONUTF8"):
            env.pop(key, None)
        result = subprocess.run(
            [sys.executable, "-c", code],
            capture_output=True,
            env=env,
            timeout=60,
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("中文标准输出", result.stdout.decode("utf-8"))
        self.assertIn("中文标准错误", result.stderr.decode("utf-8"))

    def test_estimate_row_height_uses_tallest_column(self):
        # Three short cells must not add up; the tallest column decides.
        self.assertEqual(estimate_row_height(["a", "b", "c"], [40, 40, 40]), 24)
        self.assertEqual(estimate_row_height(["a\nb"], [40]), 40)
        self.assertEqual(estimate_row_height([], []), 24)

    def test_estimate_row_height_grows_past_180_and_caps_at_409(self):
        tall = estimate_row_height(["中" * 300], [44])
        self.assertGreater(tall, 180)
        self.assertLessEqual(tall, 409)
        self.assertEqual(estimate_row_height(["中" * 5000], [44]), 409)

    def test_normalize_list_items_keeps_numeric_prefixes_without_space(self):
        self.assertEqual(
            normalize_list_items(["-1+1", "1.5 秒", "+1+1"]),
            ["-1+1", "1.5 秒", "+1+1"],
        )

    def test_normalize_list_items_strips_redundant_markers(self):
        self.assertEqual(
            normalize_list_items("1. 甲\n2. 乙\n\n- 丙"),
            ["甲", "乙", "丙"],
        )
        self.assertEqual(normalize_list_items(None), [])
        self.assertEqual(
            normalize_list_items(["甲", None, "乙"]),
            ["甲", "乙"],
        )

    def test_format_excel_multiline_numbers_multiple_items(self):
        self.assertEqual(format_excel_multiline([]), "")
        self.assertEqual(format_excel_multiline(["只有一条"]), "只有一条")
        self.assertEqual(format_excel_multiline(["甲", "乙"]), "1. 甲\n2. 乙")
        self.assertEqual(format_excel_multiline(["甲", None, "乙"]), "1. 甲\n2. 乙")

    def test_md_table_cells_escape_and_split_pipes(self):
        self.assertEqual(escape_md_table_cell("a|b"), "a\\|b")
        self.assertEqual(escape_md_table_cell(None), "")
        self.assertEqual(split_md_table_row("| a \\| b | c |"), ["a | b", "c"])

    def test_estimate_row_height_normalizes_crlf(self):
        # Usable width is 38; a trailing "\r" would push this line onto a second row.
        self.assertEqual(estimate_row_height(["a" * 38 + "\r\nx"], [40]), 40)

    def test_write_text_atomic_reports_lock_and_cleans_temp(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "out.md"
            target.write_text("old", encoding="utf-8")
            with mock.patch(
                "validation_utils.os.replace",
                side_effect=PermissionError(13, "locked", str(target)),
            ):
                with self.assertRaises(FileLockedError) as context:
                    write_text_atomic(target, "new")
            self.assertIn("文件被占用", str(context.exception))
            self.assertIn(str(target), str(context.exception))
            self.assertEqual(target.read_text(encoding="utf-8"), "old")
            temp_files = [
                path for path in Path(directory).iterdir() if path.suffix == ".tmp"
            ]
            self.assertEqual(temp_files, [])


if __name__ == "__main__":
    unittest.main()
