from __future__ import annotations

from contextlib import redirect_stderr, redirect_stdout
from copy import deepcopy
import hashlib
import io
import json
import os
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import case_baseline  # noqa: E402
from case_baseline import (  # noqa: E402
    PENDING_HEADERS,
    _merge_duplicate_entries,
    assign_missing_ids,
    canonical_tree_order,
    collect_structure_warnings,
    inspect_artifacts,
    merge_artifacts_to_ir,
    normalize_baseline,
    parse_excel_to_ir,
    parse_markdown_to_ir,
    prepare_render_ir,
    renumber_cases,
    validate_baseline,
)
from create_excel_cases import build_workbook_from_ir, save_workbook_atomic  # noqa: E402
from render_cases import build_render_ir, main as render_cases_main  # noqa: E402
from render_markdown_cases import main as render_markdown_main  # noqa: E402
from render_markdown_cases import render_markdown  # noqa: E402
from sync_artifacts import main as sync_artifacts_main  # noqa: E402
from sync_artifacts import prepare_sync_ir  # noqa: E402
import create_excel_cases  # noqa: E402
import render_markdown_cases  # noqa: E402
from traceability_matrix import (  # noqa: E402
    describe_traceability_mismatch,
    format_trace_id_list,
    parse_trace_id_list,
    sync_traceability_from_cases,
    validate_traceability_entries,
)
from validate_excel_cases import main as validate_excel_main  # noqa: E402
from validate_excel_cases import validate_excel_file  # noqa: E402
from validate_markdown_lists import fix, validate  # noqa: E402
from validate_markdown_lists import main as validate_markdown_main  # noqa: E402
import validation_utils  # noqa: E402
from validate_traceability_bundle import validate_bundle  # noqa: E402
from validation_utils import (  # noqa: E402
    FileLockedError,
    _check_targets_writable,
    commit_files_atomic,
    escape_md_table_cell,
    estimate_row_height,
    format_excel_multiline,
    is_valid_date_or_placeholder,
    normalize_list_items,
    replace_file_atomic,
    sibling_temp_path,
    split_md_table_row,
    write_text_atomic,
)


V2_MARKDOWN = """# 测试用例 - 2026-11-24 - [注册]

> 用例数量: 用例 1 条

- [ ] 注册模块

  - [ ] 功能测试

    - [ ] 正向-提交合法注册-P0

      > [TC-ZC-001] 追溯: TRACE-REG-001

      - 前置操作

        - 用户未注册。

      - 测试步骤

        - 提交注册。

      - 预期结果

        - 注册成功。

## 需求追溯

### 需求追溯表

| 追溯 ID | 来源 | 定位 | 规则 | 关联用例 |
| :--- | :--- | :--- | :--- | :--- |
| TRACE-REG-001 | 注册需求.md | 1.1 手机号 | 手机号必须为11位数字 | TC-ZC-001 |
"""


def subprocess_env():
    return {**os.environ, "PYTHONIOENCODING": "utf-8"}


def subprocess_env_without_encoding_override():
    environment = os.environ.copy()
    environment.pop("PYTHONIOENCODING", None)
    return environment


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


def valid_ir_two_cases() -> dict:
    ir = valid_ir()
    second = deepcopy(ir["cases"][0])
    second.update(
        {
            "id": "TC-ZC-002",
            "title": "[注册模块]_[注册]_[错误手机号]_[失败]",
            "polarity": "反向",
            "priority": "P1",
        }
    )
    ir["cases"].append(second)
    ir["traceability"][0]["case_ids"] = ["TC-ZC-001", "TC-ZC-002"]
    return ir


def pending_ir() -> dict:
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
    return ir


def interleaved_ir() -> dict:
    """Cases deliberately ordered 甲/乙/甲 so canonical order has to regroup them."""
    ir = valid_ir()
    base = ir["cases"][0]
    cases = []
    for index, (module, group, title) in enumerate(
        (
            ("模块甲", "分组一", "模块甲场景一"),
            ("模块乙", "分组一", "模块乙场景一"),
            ("模块甲", "分组二", "模块甲场景二"),
        ),
        start=1,
    ):
        case = deepcopy(base)
        case.update(
            {"id": f"TC-ZC-{index:03d}", "module": module, "group": group, "title": title}
        )
        cases.append(case)
    ir["cases"] = cases
    ir["traceability"] = [
        {
            "id": "TRACE-REG-001",
            "source": "注册需求.md",
            "location": "1.1 手机号",
            "rule": "手机号必须为11位数字",
            "status": "待覆盖",
            "case_ids": [],
        }
    ]
    return ir


class MarkdownV2Tests(unittest.TestCase):
    """Self-contained v2 Markdown validation with no external matrix file."""

    def _assert_error(self, text: str, expected: str, baseline=None) -> None:
        errors = validate(text, baseline)
        self.assertTrue(any(expected in error for error in errors), errors)

    def test_valid_document(self):
        self.assertEqual(validate(V2_MARKDOWN), [])

    def test_render_omits_optional_fields_and_pending_section(self):
        ir = valid_ir()
        ir["cases"][0]["test_data"] = []
        ir["cases"][0]["post_operations"] = []
        text = render_markdown(normalize_baseline(ir))
        self.assertNotIn("- 测试数据", text)
        self.assertNotIn("- 后置操作", text)
        self.assertNotIn("### 待确认项", text)
        self.assertLess(text.index("- 前置操作"), text.index("- 测试步骤"))
        self.assertLess(text.index("- 测试步骤"), text.index("- 预期结果"))
        self.assertEqual(validate(text), [])

    def test_parser_rejects_continuation_before_first_item(self):
        text = V2_MARKDOWN.replace("        - 用户未注册。", "        用户未注册。")
        self._assert_error(text, "首个列表项之前出现续行")

    def test_bundled_template_is_valid(self):
        text = (ROOT / "assets" / "md-test-case-template.md").read_text(encoding="utf-8")
        self.assertEqual(validate(text), [])

    def test_case_title_must_not_carry_ids(self):
        self._assert_error(
            V2_MARKDOWN.replace("正向-提交合法注册-P0", "正向-TC-ZC-099-提交合法注册-P0"),
            "不得包含用例 ID",
        )

    def test_checked_state_and_deep_heading_rejected(self):
        self._assert_error(
            V2_MARKDOWN.replace("- [ ] 注册模块", "- [x] 注册模块"),
            "不得使用已完成标记",
        )
        self._assert_error(
            V2_MARKDOWN.replace("  - [ ] 功能测试", "#### 深层标题\n\n  - [ ] 功能测试"),
            "不得使用四级及更深标题",
        )

    def test_duplicate_title_heading_rejected(self):
        self._assert_error(
            V2_MARKDOWN.replace(
                "# 测试用例 - 2026-11-24 - [注册]\n",
                "# 测试用例 - 2026-11-24 - [注册]\n\n"
                "# 测试用例 - 2026-11-24 - [注册]\n",
            ),
            "任务树中不得出现标题",
        )

    def test_pending_marker_rejected(self):
        self._assert_error(
            V2_MARKDOWN.replace("- 用户未注册。", "- 用户未注册。[待确认]"),
            "不得保留 [待确认] 标记",
        )

    def test_trace_section_must_be_unique(self):
        self._assert_error(V2_MARKDOWN.split("## 需求追溯")[0], "缺少唯一")
        self._assert_error(V2_MARKDOWN + "\n## 需求追溯\n", "多个 ## 需求追溯")

    def test_declared_count_must_match(self):
        self._assert_error(
            V2_MARKDOWN.replace("用例 1 条", "用例 2 条"), "用例数量声明为 2 条"
        )

    def test_field_order_and_unknown_fields(self):
        swapped = V2_MARKDOWN.replace(
            "      - 测试步骤\n\n        - 提交注册。\n\n"
            "      - 预期结果\n\n        - 注册成功。",
            "      - 预期结果\n\n        - 注册成功。\n\n"
            "      - 测试步骤\n\n        - 提交注册。",
        )
        self._assert_error(swapped, "顺序错误")
        self._assert_error(
            V2_MARKDOWN.replace("- 测试步骤", "- 备注"), "不支持字段“备注”"
        )
        self._assert_error(
            V2_MARKDOWN.replace("- 测试步骤", "- 自定义字段"), "未知的直接子字段"
        )

    def test_required_fields_must_exist(self):
        text = V2_MARKDOWN.replace("      - 测试步骤\n\n        - 提交注册。\n\n", "")
        self._assert_error(text, "缺少必填字段“测试步骤”")

    def test_reference_line_is_required_and_indented(self):
        self._assert_error(
            V2_MARKDOWN.replace("      > [TC-ZC-001] 追溯: TRACE-REG-001\n\n", ""),
            "首个非空行必须是引用行",
        )
        self._assert_error(
            V2_MARKDOWN.replace("      > [TC-ZC-001]", "        > [TC-ZC-001]"),
            "引用行缩进必须为 6 个空格",
        )

    def test_list_markers_must_match_item_count(self):
        self._assert_error(
            V2_MARKDOWN.replace("        - 提交注册。", "        1. 提交注册。"),
            "仅有一项时必须使用无序列表",
        )
        self._assert_error(
            V2_MARKDOWN.replace(
                "        - 提交注册。", "        - 提交注册。\n        - 点击保存。"
            ),
            "必须全部使用有序列表",
        )

    def test_trace_table_must_close_both_directions(self):
        self._assert_error(
            V2_MARKDOWN.replace("| TC-ZC-001 |\n", "| TC-ZC-002 |\n"),
            "关联用例与正文引用不一致",
        )
        self._assert_error(
            V2_MARKDOWN.replace("追溯: TRACE-REG-001", "追溯: TRACE-NOPE"),
            "引用了不存在的追溯 ID",
        )

    def test_trace_table_header_must_match(self):
        self._assert_error(
            V2_MARKDOWN.replace("| 关联用例 |", "| 关联用例 ID |"),
            "需求追溯表头必须为",
        )

    def test_pending_section_rules(self):
        pending = (
            "\n### 待确认项\n\n"
            "| 追溯 ID | 来源 | 定位 | 未明确项 | 当前测试假设 | 确认问题 |\n"
            "| :--- | :--- | :--- | :--- | :--- | :--- |\n"
            "| TRACE-REG-Q01 | 注册需求.md | 1.2 超时 | 超时行为未定义 | 暂按失败处理 | 超时是否重试？ |\n"
        )
        self.assertEqual(validate(V2_MARKDOWN + pending), [])
        self._assert_error(
            V2_MARKDOWN
            + pending.replace("| :--- | :--- | :--- | :--- | :--- | :--- |\n", ""),
            "缺少标准分隔行",
        )
        self._assert_error(
            V2_MARKDOWN + pending.split("| TRACE-REG-Q01")[0],
            "待确认项表格不得为空",
        )
        self._assert_error(
            V2_MARKDOWN + pending.replace("TRACE-REG-Q01", "TRACE-REG-001"),
            "追溯 ID 重复",
        )
        self._assert_error(
            V2_MARKDOWN.replace("追溯: TRACE-REG-001", "追溯: TRACE-REG-Q01") + pending,
            "不得引用待确认条目",
        )
        self._assert_error(
            V2_MARKDOWN + "\n### 未知小节\n\n| a | b |\n| :--- | :--- |\n",
            "未知三级标题",
        )

    def test_fix_renumbers_multi_item_lists(self):
        text = (
            "- 测试步骤\r\n\r\n  - 第一步\r\n  - 第二步\r\n\r\n"
            "## 需求追溯\r\n\r\n- 测试步骤\r\n  - 甲\r\n  - 乙\r\n"
        )
        fixed = fix(text)
        self.assertIn("  1. 第一步\r\n  2. 第二步", fixed)
        self.assertTrue(fixed.endswith("- 测试步骤\r\n  - 甲\r\n  - 乙\r\n"))
        self.assertNotIn("1. 甲", fixed)

    def test_fix_covers_all_five_fields_and_skips_trace_section(self):
        fields = ["前置操作", "测试数据", "测试步骤", "预期结果", "后置操作"]
        blocks = []
        for field_index, field in enumerate(fields):
            first, second = f"字段{field_index}甲", f"字段{field_index}乙"
            blocks.append(f"        - {field}\n          - {first}\n          - {second}\n")
        text = "\n".join(blocks) + (
            "\n## 需求追溯\n\n- 测试步骤\n  - 表内甲\n  - 表内乙\n"
        )

        fixed = fix(text)

        for field_index, field in enumerate(fields):
            self.assertIn(f"        - {field}\n", fixed)
            self.assertIn(
                f"          1. 字段{field_index}甲\n          2. 字段{field_index}乙",
                fixed,
            )
        self.assertTrue(fixed.endswith("- 测试步骤\n  - 表内甲\n  - 表内乙\n"))
        self.assertNotIn("1. 表内甲", fixed)

    def test_baseline_comparison_ignores_design_method(self):
        errors = validate(V2_MARKDOWN, valid_ir())
        self.assertTrue(any("title 不一致" in error for error in errors), errors)
        self.assertFalse(any("design_method" in error for error in errors), errors)

    def test_cli_validates_and_fixes(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cases.md"
            path.write_text(
                V2_MARKDOWN.replace(
                    "        - 提交注册。", "        - 提交注册。\n        - 点击保存。"
                ),
                encoding="utf-8",
            )
            env = subprocess_env()
            script = str(ROOT / "scripts" / "validate_markdown_lists.py")
            failed = subprocess.run(
                [sys.executable, script, str(path)],
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=env,
            )
            self.assertEqual(failed.returncode, 1)
            self.assertIn("必须全部使用有序列表", failed.stderr)
            fixed = subprocess.run(
                [sys.executable, script, str(path), "--fix"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=env,
            )
            self.assertEqual(fixed.returncode, 0, fixed.stderr)
            self.assertIn("Markdown 校验通过", fixed.stdout)
            content = path.read_text(encoding="utf-8")
            self.assertIn("        1. 提交注册。", content)
            self.assertIn("        2. 点击保存。", content)
            self.assertEqual(
                [item.name for item in path.parent.iterdir()], [path.name]
            )

    def test_cli_fix_preserves_crlf(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cases.md"
            path.write_bytes(
                V2_MARKDOWN.replace(
                    "        - 提交注册。", "        - 提交注册。\n        - 点击保存。"
                )
                .replace("\n", "\r\n")
                .encode("utf-8")
            )
            self.assertEqual(validate_markdown_main([str(path), "--fix"]), 0)
            data = path.read_bytes()
            self.assertNotIn(b"\n", data.replace(b"\r\n", b""))
            content = data.decode("utf-8")
            self.assertIn("        1. 提交注册。\r\n", content)
            self.assertIn("        2. 点击保存。\r\n", content)

    def test_cli_missing_file(self):
        self.assertEqual(validate_markdown_main([str(ROOT / "不存在.md")]), 1)


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


class ExcelV2Tests(unittest.TestCase):
    """Self-contained v2 Excel validation, including sheet layout rules."""

    def _assert_error(self, mutate, expected: str) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cases.xlsx"
            save_workbook_atomic(build_workbook_from_ir(valid_ir()), path)
            workbook = load_workbook(path)
            mutate(workbook)
            workbook.save(path)
            errors = validate_excel_file(path)
        self.assertTrue(any(expected in error for error in errors), errors)

    def test_valid_workbook_passes(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cases.xlsx"
            save_workbook_atomic(build_workbook_from_ir(valid_ir()), path)
            self.assertEqual(validate_excel_file(path), [])
            self.assertEqual(validate_excel_main([str(path)]), 0)

    def test_bundled_template_matches_validator(self):
        path = ROOT / "assets" / "excel-test-case-template.xlsx"
        self.assertEqual(validate_excel_file(path), [])
        workbook = load_workbook(path)
        self.assertEqual(workbook.sheetnames, ["测试用例", "生成说明", "需求追溯"])
        self.assertEqual(workbook["测试用例"]["B3"].value, "TC-SLXQ-001")
        self.assertEqual(workbook["测试用例"]["B4"].value, "TC-SLXQ-002")
        self.assertEqual(workbook["生成说明"]["B4"].value, "2")

    def test_sheet_layout_rules(self):
        self._assert_error(
            lambda workbook: workbook["测试用例"].__setitem__("B2", "用例编号"),
            "无法识别的工作表表头",
        )
        self._assert_error(
            lambda workbook: setattr(workbook["测试用例"], "freeze_panes", None),
            "必须冻结窗格至 A3",
        )
        self._assert_error(
            lambda workbook: setattr(
                workbook["测试用例"].auto_filter, "ref", "A2:K2"
            ),
            "筛选范围必须为 A2:K3",
        )
        self._assert_error(
            lambda workbook: workbook["测试用例"].merge_cells("A2:B2"),
            "必须恰好包含 A1:K1 合并标题区域",
        )

    def test_extra_column_rules(self):
        self._assert_error(
            lambda workbook: workbook["测试用例"].__setitem__("L3", "越界"),
            "只允许 A–K 共 11 列",
        )
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cases.xlsx"
            save_workbook_atomic(build_workbook_from_ir(valid_ir()), path)
            workbook = load_workbook(path)
            workbook["测试用例"]["L3"] = " "
            workbook.save(path)
            self.assertEqual(validate_excel_file(path), [])

    def test_required_columns_and_enums(self):
        self._assert_error(
            lambda workbook: workbook["测试用例"].__setitem__("I3", None),
            "缺少必填字段“测试步骤”",
        )
        self._assert_error(
            lambda workbook: workbook["测试用例"].__setitem__("F3", "P9"), "priority"
        )
        self._assert_error(
            lambda workbook: workbook["测试用例"].__setitem__("E3", "随意方法"),
            "design_method",
        )
        self._assert_error(
            lambda workbook: workbook["测试用例"].__setitem__("D3", "异常测试"), "type"
        )

    def test_formula_cells_are_rejected(self):
        self._assert_error(
            lambda workbook: workbook["测试用例"].__setitem__("G3", "=1+1"),
            "包含公式",
        )
        self._assert_error(
            lambda workbook: workbook["需求追溯"].__setitem__("E3", "=1+1"),
            "包含公式",
        )
        self._assert_error(
            lambda workbook: workbook["生成说明"].__setitem__("B2", "=1+1"),
            "包含公式",
        )

    def test_info_sheet_rules(self):
        self._assert_error(
            lambda workbook: workbook["生成说明"].__setitem__("B2", ""), "不能为空"
        )
        self._assert_error(
            lambda workbook: workbook["生成说明"].__setitem__("B3", "2026-02-30"),
            "上线日期",
        )
        self._assert_error(
            lambda workbook: workbook["生成说明"].__setitem__("B4", "9"),
            "与实际用例行数 1 不一致",
        )
        self._assert_error(
            lambda workbook: workbook["生成说明"].delete_rows(6),
            "缺少字段“补充说明”",
        )
        self._assert_error(
            lambda workbook: workbook["生成说明"].__setitem__("C2", "多余"),
            "只允许 A–B 两列",
        )

    def test_traceability_sheet_rules(self):
        self._assert_error(
            lambda workbook: workbook["需求追溯"].delete_rows(3),
            "至少需要 1 行有效数据",
        )
        self._assert_error(
            lambda workbook: workbook["需求追溯"].__setitem__("E2", "关联用例"),
            "表头必须为",
        )
        self._assert_error(
            lambda workbook: workbook["需求追溯"].__setitem__("F3", "越界"),
            "只允许 A–E 共 5 列",
        )

    def test_pending_sheet_rules(self):
        def add_pending(workbook, row_values):
            sheet = workbook.create_sheet("待确认项")
            sheet["A1"] = "待确认项清单"
            sheet.merge_cells("A1:F1")
            for index, header in enumerate(PENDING_HEADERS, start=1):
                sheet.cell(row=2, column=index, value=header)
            for index, value in enumerate(row_values, start=1):
                sheet.cell(row=3, column=index, value=value)
            sheet.freeze_panes = "A3"
            sheet.auto_filter.ref = "A2:F3"

        row = [
            "TRACE-REG-Q01",
            "注册需求.md",
            "1.2 超时",
            "超时行为未定义",
            "暂按失败处理",
            "超时是否重试？",
        ]
        self._assert_error(
            lambda workbook: add_pending(workbook, [None] * 6), "禁止空壳工作表"
        )
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cases.xlsx"
            save_workbook_atomic(build_workbook_from_ir(valid_ir()), path)
            workbook = load_workbook(path)
            add_pending(workbook, row)
            workbook.save(path)
            self.assertEqual(validate_excel_file(path), [])

    def test_unexpected_sheet_rejected(self):
        self._assert_error(
            lambda workbook: workbook.create_sheet("Sheet"), "多余的工作表"
        )

    def test_main_sheet_links_must_match_trace_sheet(self):
        self._assert_error(
            lambda workbook: workbook["需求追溯"].__setitem__("E3", "TC-ZC-999"),
            "关联用例与正文引用不一致",
        )

    def test_missing_file_and_reserved_sheet_name(self):
        self.assertTrue(validate_excel_file(ROOT / "不存在.xlsx"))
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cases.xlsx"
            save_workbook_atomic(build_workbook_from_ir(valid_ir()), path)
            self.assertTrue(
                any("保留名" in error for error in validate_excel_file(path, sheet_name="需求追溯"))
            )
            self.assertTrue(
                any("缺少名为" in error for error in validate_excel_file(path, sheet_name="不存在"))
            )


class RoundTripAndSyncTests(unittest.TestCase):
    """Render, cross-validate and sync the two formats as one delivery."""

    def _write_ir(self, path: Path, ir: dict) -> None:
        path.write_text(json.dumps(ir, ensure_ascii=False), encoding="utf-8")

    def _render(self, directory: Path, ir: dict, extra: tuple[str, ...] = ()):
        baseline = directory / "ir.json"
        markdown = directory / "cases.md"
        excel = directory / "cases.xlsx"
        self._write_ir(baseline, ir)
        code = render_cases_main(
            [
                "--baseline",
                str(baseline),
                "--markdown",
                str(markdown),
                "--excel",
                str(excel),
                *extra,
            ]
        )
        return code, baseline, markdown, excel

    def _sync(self, markdown: Path, excel: Path | None, extra: tuple[str, ...] = ()):
        argv = ["--from-markdown", str(markdown)]
        if excel is not None:
            argv.extend(["--excel", str(excel)])
        return sync_artifacts_main([*argv, *extra])

    def _case_block(self, lines: list[str], marker: str) -> tuple[int, int]:
        start = next(index for index, line in enumerate(lines) if marker in line)
        while start > 0 and not re.match(
            r"^ *- \[ \] (?:正向|反向)-.+-P[0-4]\s*$", lines[start]
        ):
            start -= 1
        indent = len(lines[start]) - len(lines[start].lstrip())
        end = start + 1
        while end < len(lines):
            current = lines[end]
            if current.strip() and len(current) - len(current.lstrip()) <= indent:
                break
            end += 1
        return start, end

    def test_render_cases_writes_both_formats(self):
        with tempfile.TemporaryDirectory() as directory:
            code, baseline, markdown, excel = self._render(Path(directory), valid_ir())
            self.assertEqual(code, 0)
            self.assertTrue(baseline.exists())
            self.assertEqual(validate(markdown.read_text(encoding="utf-8")), [])
            self.assertEqual(validate_excel_file(excel), [])

    def test_render_cases_assigns_ids_and_cleans_baseline(self):
        ir = valid_ir()
        ir["cases"][0]["id"] = "TC-NEW"
        ir["traceability"][0].update({"status": "待覆盖", "case_ids": []})
        with tempfile.TemporaryDirectory() as directory:
            code, baseline, markdown, excel = self._render(
                Path(directory),
                ir,
                ("--assign-ids", "--id-prefix", "TC-ZC", "--cleanup-baseline"),
            )
            self.assertEqual(code, 0)
            self.assertFalse(baseline.exists())
            self.assertIn("[TC-ZC-001]", markdown.read_text(encoding="utf-8"))
            self.assertEqual(validate_excel_file(excel), [])

    def test_render_cases_blocks_uncovered_entries(self):
        ir = valid_ir()
        ir["traceability"].append(
            {
                "id": "TRACE-REG-002",
                "source": "注册需求.md",
                "location": "1.3 验证码",
                "rule": "验证码 5 分钟内有效",
                "status": "待覆盖",
                "case_ids": [],
            }
        )
        with tempfile.TemporaryDirectory() as directory:
            code, _, markdown, excel = self._render(Path(directory), ir)
            self.assertEqual(code, 1)
            self.assertFalse(markdown.exists())
            self.assertFalse(excel.exists())

    def test_render_cases_requires_force_to_overwrite(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            markdown = directory_path / "cases.md"
            markdown.write_text("旧文件", encoding="utf-8")
            code, _, _, _ = self._render(directory_path, valid_ir())
            self.assertEqual(code, 1)
            self.assertEqual(markdown.read_text(encoding="utf-8"), "旧文件")

    def test_render_cases_rejects_missing_design_method(self):
        ir = valid_ir()
        ir["cases"][0]["design_method"] = ""
        with tempfile.TemporaryDirectory() as directory:
            code, _, markdown, excel = self._render(Path(directory), ir)
            self.assertEqual(code, 1)
            self.assertFalse(markdown.exists())
            self.assertFalse(excel.exists())

    def test_render_ir_orders_cases_before_syncing_links(self):
        ir, errors, _ = build_render_ir(interleaved_ir())
        self.assertEqual(errors, [])
        self.assertEqual(
            [case["id"] for case in ir["cases"]],
            ["TC-ZC-001", "TC-ZC-003", "TC-ZC-002"],
        )
        self.assertEqual(
            ir["traceability"][0]["case_ids"],
            ["TC-ZC-001", "TC-ZC-003", "TC-ZC-002"],
        )
        self.assertEqual(validate_baseline(ir), [])

    def test_sync_assigns_ids_and_backfills_design_method(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            markdown = directory_path / "cases.md"
            excel = directory_path / "cases.xlsx"
            markdown.write_text(
                render_markdown(normalize_baseline(valid_ir_two_cases())),
                encoding="utf-8",
            )
            save_workbook_atomic(build_workbook_from_ir(valid_ir()), excel)
            markdown.write_text(
                markdown.read_text(encoding="utf-8").replace(
                    "[TC-ZC-002]", "[TC-NEW]"
                ),
                encoding="utf-8",
            )
            code = self._sync(
                markdown,
                excel,
                (
                    "--assign-ids",
                    "--id-prefix",
                    "TC-ZC",
                    "--default-design-method",
                    "等价类划分",
                ),
            )
            self.assertEqual(code, 0)
            self.assertIn("[TC-ZC-002]", markdown.read_text(encoding="utf-8"))
            self.assertEqual(validate_excel_file(excel), [])
            sheet = load_workbook(excel)["测试用例"]
            self.assertEqual(sheet["B4"].value, "TC-ZC-002")
            self.assertEqual(sheet["E4"].value, "等价类划分")

    def test_sync_blocks_missing_design_method_without_default(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            markdown = directory_path / "cases.md"
            excel = directory_path / "cases.xlsx"
            markdown.write_text(
                render_markdown(normalize_baseline(valid_ir_two_cases())).replace(
                    "[TC-ZC-002]", "[TC-NEW]"
                ),
                encoding="utf-8",
            )
            save_workbook_atomic(build_workbook_from_ir(valid_ir()), excel)
            code = self._sync(
                markdown, excel, ("--assign-ids", "--id-prefix", "TC-ZC")
            )
            self.assertEqual(code, 1)
            self.assertIn("[TC-NEW]", markdown.read_text(encoding="utf-8"))

    def test_sync_blocks_lost_traceability(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            markdown = directory_path / "cases.md"
            excel = directory_path / "cases.xlsx"
            _, tail = V2_MARKDOWN.split("## 需求追溯\n", 1)
            text = (
                "# 测试用例 - 2026-11-24 - [注册]\n\n"
                "> 用例数量: 用例 0 条\n\n"
                "## 需求追溯\n"
                + tail
            )
            markdown.write_text(text, encoding="utf-8")
            save_workbook_atomic(build_workbook_from_ir(valid_ir()), excel)
            ir, errors, _ = prepare_sync_ir(markdown, excel)
            self.assertEqual(ir, {})
            self.assertTrue(
                any("失去全部关联用例" in error for error in errors), errors
            )
            self.assertEqual(self._sync(markdown, excel), 1)
            self.assertEqual(markdown.read_text(encoding="utf-8"), text)

    def test_sync_renumbers_ids(self):
        ir = valid_ir_two_cases()
        ir["cases"][0]["id"] = "TC-ZC-007"
        ir["cases"][1]["id"] = "TC-ZC-009"
        ir["traceability"][0]["case_ids"] = ["TC-ZC-007", "TC-ZC-009"]
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            markdown = directory_path / "cases.md"
            excel = directory_path / "cases.xlsx"
            markdown.write_text(render_markdown(normalize_baseline(ir)), encoding="utf-8")
            save_workbook_atomic(build_workbook_from_ir(ir), excel)
            code = self._sync(markdown, excel, ("--renumber", "--id-prefix", "TC-ZC"))
            self.assertEqual(code, 0)
            content = markdown.read_text(encoding="utf-8")
            self.assertIn("[TC-ZC-001]", content)
            self.assertIn("[TC-ZC-002]", content)
            self.assertNotIn("TC-ZC-007", content)
            sheet = load_workbook(excel)["测试用例"]
            self.assertEqual([sheet["B3"].value, sheet["B4"].value], ["TC-ZC-001", "TC-ZC-002"])

    def test_sync_keeps_existing_excel_meta(self):
        ir = valid_ir()
        ir["meta"]["coverage"] = "覆盖注册主流程"
        ir["meta"]["notes"] = "基于需求文档生成"
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            code, _, markdown, excel = self._render(directory_path, ir)
            self.assertEqual(code, 0)
            self.assertEqual(self._sync(markdown, excel), 0)
            info = load_workbook(excel)["生成说明"]
            values = {
                info.cell(row=row, column=1).value: info.cell(row=row, column=2).value
                for row in range(2, 7)
            }
        self.assertEqual(values["覆盖范围"], "覆盖注册主流程")
        self.assertEqual(values["补充说明"], "基于需求文档生成")

    def test_sync_markdown_only(self):
        with tempfile.TemporaryDirectory() as directory:
            markdown = Path(directory) / "cases.md"
            markdown.write_text(
                render_markdown(normalize_baseline(valid_ir_two_cases())),
                encoding="utf-8",
            )
            self.assertEqual(self._sync(markdown, None), 0)
            self.assertEqual(validate(markdown.read_text(encoding="utf-8")), [])

    def test_export_ir_from_both_formats_preserves_pending(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            code, _, markdown, excel = self._render(directory_path, pending_ir())
            self.assertEqual(code, 0)
            exported = directory_path / "exported.json"
            self.assertEqual(
                case_baseline.main(
                    [
                        "export",
                        "--markdown",
                        str(markdown),
                        "--excel",
                        str(excel),
                        "--output",
                        str(exported),
                    ]
                ),
                0,
            )
            data = json.loads(exported.read_text(encoding="utf-8"))
            self.assertEqual([case["id"] for case in data["cases"]], ["TC-ZC-001"])
            self.assertEqual(data["cases"][0]["design_method"], "等价类划分")
            self.assertEqual(data["meta"]["coverage"], "覆盖注册主流程")
            pending = [
                entry for entry in data["traceability"] if entry["status"] == "待确认"
            ]
            self.assertEqual(pending[0]["question"], "超时是否重试？")

            data["traceability"] = [
                entry for entry in data["traceability"] if entry["status"] != "待确认"
            ]
            baseline = directory_path / "cleared.json"
            baseline.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
            cleared_md = directory_path / "cleared.md"
            cleared_xlsx = directory_path / "cleared.xlsx"
            self.assertEqual(
                render_cases_main(
                    [
                        "--baseline",
                        str(baseline),
                        "--markdown",
                        str(cleared_md),
                        "--excel",
                        str(cleared_xlsx),
                    ]
                ),
                0,
            )
            self.assertEqual(
                load_workbook(cleared_xlsx).sheetnames,
                ["测试用例", "生成说明", "需求追溯"],
            )
            self.assertNotIn(
                "### 待确认项", cleared_md.read_text(encoding="utf-8")
            )
            cleared_out = directory_path / "cleared_out.json"
            self.assertEqual(
                case_baseline.main(
                    [
                        "export",
                        "--markdown",
                        str(cleared_md),
                        "--excel",
                        str(cleared_xlsx),
                        "--output",
                        str(cleared_out),
                    ]
                ),
                0,
            )
            cleared = json.loads(cleared_out.read_text(encoding="utf-8"))
            self.assertFalse(
                any(entry["status"] == "待确认" for entry in cleared["traceability"])
            )

    def test_sync_manual_edit_preserves_gap_and_backfills_new_case(self):
        ir = valid_ir()
        first = ir["cases"][0]
        third = deepcopy(first)
        third.update(
            {
                "id": "TC-ZC-003",
                "title": "[注册模块]_[注册]_[边界手机号]_[成功]",
                "design_method": "边界值分析",
            }
        )
        fifth = deepcopy(first)
        fifth.update(
            {
                "id": "TC-ZC-005",
                "title": "[注册模块]_[注册]_[异常手机号]_[失败]",
                "polarity": "反向",
                "design_method": "错误推测法",
            }
        )
        ir["cases"] = [first, third, fifth]
        ir["traceability"][0]["case_ids"] = ["TC-ZC-001", "TC-ZC-003", "TC-ZC-005"]

        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            code, _, markdown, excel = self._render(directory_path, ir)
            self.assertEqual(code, 0)

            lines = markdown.read_text(encoding="utf-8").splitlines()
            lines = [
                line.replace(
                    "[注册模块]_[注册]_[合法手机号]_[成功]",
                    "[注册模块]_[注册]_[合法手机号]_[提交成功]",
                )
                for line in lines
            ]
            start, end = self._case_block(lines, "[TC-ZC-003]")
            del lines[start:end]

            start, end = self._case_block(lines, "[TC-ZC-001]")
            clone = list(lines[start:end])
            clone[0] = clone[0].replace(
                "正向-[注册模块]_[注册]_[合法手机号]_[提交成功]-P0",
                "反向-新增无号场景-P2",
            )
            clone = [
                re.sub(
                    r"^(\s*)> \[TC-ZC-001\] 追溯:",
                    r"\1> 追溯:",
                    line,
                )
                for line in clone
            ]
            insert_at, _ = self._case_block(lines, "[TC-ZC-005]")
            lines[insert_at:insert_at] = clone
            markdown.write_text("\n".join(lines) + "\n", encoding="utf-8")
            before = markdown.read_text(encoding="utf-8")

            self.assertEqual(
                self._sync(
                    markdown, excel, ("--assign-ids", "--id-prefix", "TC-ZC")
                ),
                1,
            )
            self.assertEqual(markdown.read_text(encoding="utf-8"), before)
            self.assertNotIn("TC-ZC-006", before)

            self.assertEqual(
                self._sync(
                    markdown,
                    excel,
                    (
                        "--assign-ids",
                        "--id-prefix",
                        "TC-ZC",
                        "--default-design-method",
                        "等价类划分",
                    ),
                ),
                0,
            )
            content = markdown.read_text(encoding="utf-8")
            self.assertIn("[TC-ZC-006]", content)
            self.assertNotIn("TC-ZC-003", content)
            self.assertIn("反向-新增无号场景-P2", content)
            self.assertEqual(validate(content), [])
            self.assertEqual(validate_excel_file(excel), [])
            sheet = load_workbook(excel)["测试用例"]
            self.assertEqual(
                [sheet.cell(row, 2).value for row in range(3, 6)],
                ["TC-ZC-001", "TC-ZC-006", "TC-ZC-005"],
            )
            self.assertEqual(
                [sheet.cell(row, 5).value for row in range(3, 6)],
                ["等价类划分", "等价类划分", "错误推测法"],
            )
            self.assertEqual(sheet["C3"].value, "[注册模块]_[注册]_[合法手机号]_[提交成功]")

    def test_sync_first_time_excel_requires_default_design_method(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            markdown = directory_path / "cases.md"
            excel = directory_path / "cases.xlsx"
            markdown.write_text(
                render_markdown(normalize_baseline(valid_ir())).replace(
                    "[TC-ZC-001]", "[TC-NEW]"
                ),
                encoding="utf-8",
            )
            before = markdown.read_text(encoding="utf-8")
            self.assertEqual(
                self._sync(
                    markdown, excel, ("--assign-ids", "--id-prefix", "TC-ZC")
                ),
                1,
            )
            self.assertFalse(excel.exists())
            self.assertEqual(markdown.read_text(encoding="utf-8"), before)

            self.assertEqual(
                self._sync(
                    markdown,
                    excel,
                    (
                        "--assign-ids",
                        "--id-prefix",
                        "TC-ZC",
                        "--default-design-method",
                        "等价类划分",
                    ),
                ),
                0,
            )
            self.assertTrue(excel.exists())
            self.assertEqual(validate_excel_file(excel), [])
            self.assertEqual(load_workbook(excel)["测试用例"]["E3"].value, "等价类划分")

    def test_sync_write_failure_rolls_back_both_files(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            code, _, markdown, excel = self._render(directory_path, valid_ir())
            self.assertEqual(code, 0)
            before_markdown = markdown.read_bytes()
            before_excel = excel.read_bytes()
            original_replace = validation_utils.replace_file_atomic
            calls = {"count": 0}

            def flaky_replace(source, target):
                calls["count"] += 1
                if calls["count"] == 2:
                    raise PermissionError(13, "locked", str(target))
                return original_replace(source, target)

            with mock.patch.object(
                validation_utils, "replace_file_atomic", side_effect=flaky_replace
            ):
                self.assertEqual(self._sync(markdown, excel), 1)

            self.assertEqual(markdown.read_bytes(), before_markdown)
            self.assertEqual(excel.read_bytes(), before_excel)
            self.assertEqual(
                sorted(path.name for path in directory_path.iterdir()),
                ["cases.md", "cases.xlsx", "ir.json"],
            )

    def test_bundle_detects_id_trace_and_pending_mismatch(self):
        mutations = (
            ("TC-ZC-001", "TC-ZC-099", "TC-ZC-099"),
            ("手机号必须为11位数字", "手机号必须为10位数字", "rule"),
            ("超时是否重试？", "超时是否重试或失败？", "question"),
        )
        for old, new, marker in mutations:
            with self.subTest(marker=marker):
                with tempfile.TemporaryDirectory() as directory:
                    directory_path = Path(directory)
                    code, _, markdown, excel = self._render(
                        directory_path, pending_ir()
                    )
                    self.assertEqual(code, 0)
                    markdown.write_text(
                        markdown.read_text(encoding="utf-8").replace(old, new),
                        encoding="utf-8",
                    )
                    errors = validate_bundle(
                        markdown_path=markdown, excel_path=excel
                    )
                self.assertTrue(errors)
                if marker == "TC-ZC-099":
                    self.assertTrue(
                        any(marker in error for error in errors), errors
                    )
                elif marker == "rule":
                    self.assertTrue(
                        any(
                            "TRACE-REG-001" in error and "rule" in error
                            for error in errors
                        ),
                        errors,
                    )
                else:
                    self.assertTrue(
                        any(
                            "TRACE-REG-Q01" in error and "question" in error
                            for error in errors
                        ),
                        errors,
                    )

    def test_markdown_only_render_without_openpyxl(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            blocker = directory_path / "blocker"
            blocker.mkdir()
            (blocker / "sitecustomize.py").write_text(
                "import sys\n"
                "\n"
                "class _BlockOpenpyxl:\n"
                "    def find_spec(self, fullname, path=None, target=None):\n"
                "        if fullname == 'openpyxl' or fullname.startswith('openpyxl.'):\n"
                "            raise ImportError('openpyxl blocked for markdown-only test')\n"
                "        return None\n"
                "\n"
                "sys.meta_path.insert(0, _BlockOpenpyxl())\n",
                encoding="utf-8",
            )
            baseline = directory_path / "ir.json"
            markdown = directory_path / "cases.md"
            exported = directory_path / "exported.json"
            baseline.write_text(
                json.dumps(valid_ir(), ensure_ascii=False), encoding="utf-8"
            )
            env = {**subprocess_env(), "PYTHONPATH": str(blocker)}

            def run(*args: str) -> subprocess.CompletedProcess:
                return subprocess.run(
                    [sys.executable, *args],
                    cwd=directory_path,
                    env=env,
                    text=True,
                    capture_output=True,
                )

            blocked = run("-c", "import openpyxl")
            self.assertNotEqual(blocked.returncode, 0)
            self.assertIn("blocked", blocked.stderr)

            commands = (
                (
                    str(ROOT / "scripts" / "render_cases.py"),
                    "--baseline",
                    str(baseline),
                    "--markdown",
                    str(markdown),
                ),
                (str(ROOT / "scripts" / "validate_markdown_lists.py"), str(markdown)),
                (
                    str(ROOT / "scripts" / "validate_traceability_bundle.py"),
                    "--markdown",
                    str(markdown),
                ),
                (
                    str(ROOT / "scripts" / "case_baseline.py"),
                    "export",
                    "--markdown",
                    str(markdown),
                    "--output",
                    str(exported),
                ),
                (
                    str(ROOT / "scripts" / "sync_artifacts.py"),
                    "--from-markdown",
                    str(markdown),
                ),
            )
            for command in commands:
                with self.subTest(script=Path(command[0]).name):
                    result = run(*command)
                    self.assertEqual(
                        result.returncode,
                        0,
                        f"{Path(command[0]).name} 失败:\n{result.stdout}\n{result.stderr}",
                    )
            self.assertEqual(validate(markdown.read_text(encoding="utf-8")), [])
            self.assertTrue(exported.is_file())

    def test_bundle_reports_cross_format_field_diff(self):
        with tempfile.TemporaryDirectory() as directory:
            code, _, markdown, excel = self._render(Path(directory), valid_ir())
            self.assertEqual(code, 0)
            markdown.write_text(
                markdown.read_text(encoding="utf-8").replace(
                    "[合法手机号]", "[非法手机号]"
                ),
                encoding="utf-8",
            )
            errors = validate_bundle(markdown_path=markdown, excel_path=excel)
        self.assertTrue(any("title 不一致" in error for error in errors), errors)

    def test_bundle_coverage_only_skips_layout(self):
        with tempfile.TemporaryDirectory() as directory:
            code, _, markdown, excel = self._render(Path(directory), valid_ir())
            self.assertEqual(code, 0)
            workbook = load_workbook(excel)
            workbook["测试用例"].freeze_panes = None
            workbook.save(excel)
            self.assertTrue(validate_bundle(markdown_path=markdown, excel_path=excel))
            coverage = validate_bundle(
                markdown_path=markdown, excel_path=excel, coverage_only=True
            )
        self.assertEqual(coverage, [])

    def test_bundle_requires_a_file(self):
        self.assertEqual(validate_bundle(), ["必须提供 Markdown 或 Excel 至少一个文件。"])


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

    def test_assign_ids_infers_prefix_and_keeps_ids_unique(self):
        cases = [
            {"id": "TC-A-001"},
            {"id": "TC-A-002"},
            {"id": "TC-B-001"},
            {"id": "TC-NEW"},
            {"id": "NEW"},
        ]
        assigned = assign_missing_ids(cases)
        self.assertEqual(assigned, ["TC-A-003", "TC-A-004"])
        self.assertEqual(len({case["id"] for case in cases}), len(cases))

    def test_renumber_multi_prefix_syncs_links(self):
        cases = [
            {"id": "TC-A-010", "module": "M", "type": "功能测试", "group": ""},
            {"id": "TC-B-020", "module": "M", "type": "功能测试", "group": ""},
        ]
        traceability = [
            {
                "id": "TRACE-1",
                "status": "已覆盖",
                "case_ids": ["TC-A-010", "TC-B-020"],
            }
        ]
        mapping = renumber_cases(cases, traceability)
        self.assertEqual([case["id"] for case in cases], ["TC-A-001", "TC-B-001"])
        self.assertEqual(mapping["TC-B-020"], "TC-B-001")
        self.assertEqual(traceability[0]["case_ids"], ["TC-A-001", "TC-B-001"])

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
            (base.replace("- 测试数据", "- 备注"), "不支持字段"),
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
            ):
                self.assertEqual(create_excel_cases.main(argv[1:]), 1)


def bracketed_ir() -> dict:
    """Same module, titles whose first bracket differs per case."""
    ir = valid_ir()
    base = ir["cases"][0]
    cases = []
    for index, title in enumerate(("[登录]_场景1", "[风控]_场景2", "[登录]_场景3"), start=1):
        case = deepcopy(base)
        case.update(
            {"id": f"TC-CS-{index:03d}", "module": "用户中心", "group": "", "title": title}
        )
        case["traceability_refs"] = ["TRACE-1"]
        cases.append(case)
    ir["cases"] = cases
    ir["traceability"] = [
        {
            "id": "TRACE-1",
            "source": "用户需求.md",
            "location": "1.1 登录",
            "rule": "登录成功",
            "status": "已覆盖",
            "case_ids": ["TC-CS-001", "TC-CS-002", "TC-CS-003"],
        }
    ]
    return ir


class ReviewFixTests(unittest.TestCase):
    """Regressions for the blocking review items."""

    def _assert_markdown_error(self, text: str, expected: str) -> None:
        errors = validate(text)
        self.assertTrue(any(expected in error for error in errors), errors)

    def _assert_excel_error(self, mutate, expected: str, ir=None) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cases.xlsx"
            save_workbook_atomic(build_workbook_from_ir(ir or valid_ir()), path)
            workbook = load_workbook(path)
            mutate(workbook)
            workbook.save(path)
            errors = validate_excel_file(path)
        self.assertTrue(any(expected in error for error in errors), errors)

    def _write_ir(self, path: Path, ir: dict) -> None:
        path.write_text(json.dumps(ir, ensure_ascii=False), encoding="utf-8")

    def _render(self, directory: Path, ir: dict):
        baseline = directory / "ir.json"
        markdown = directory / "cases.md"
        excel = directory / "cases.xlsx"
        self._write_ir(baseline, ir)
        code = render_cases_main(
            [
                "--baseline",
                str(baseline),
                "--markdown",
                str(markdown),
                "--excel",
                str(excel),
            ]
        )
        return code, baseline, markdown, excel

    def test_count_declaration_must_exist_and_be_unique(self):
        self._assert_markdown_error(
            V2_MARKDOWN.replace("> 用例数量: 用例 1 条\n\n", ""),
            "缺少“> 用例数量",
        )
        duplicated = V2_MARKDOWN.replace(
            "> 用例数量: 用例 1 条",
            "> 用例数量: 用例 1 条\n\n> 用例数量: 用例 1 条",
        )
        self._assert_markdown_error(duplicated, "重复出现")

    def test_reference_order_must_match_trace_table(self):
        text = V2_MARKDOWN.replace(
            "| TRACE-REG-001 | 注册需求.md | 1.1 手机号 | 手机号必须为11位数字 | TC-ZC-001 |",
            "| TRACE-REG-001 | 注册需求.md | 1.1 手机号 | 手机号必须为11位数字 | TC-ZC-001 |\n"
            "| TRACE-REG-002 | 注册需求.md | 1.2 验证码 | 验证码 5 分钟内有效 | TC-ZC-001 |",
        )
        ordered = text.replace(
            "追溯: TRACE-REG-001", "追溯: TRACE-REG-001, TRACE-REG-002"
        )
        self.assertEqual(validate(ordered), [])
        reversed_text = text.replace(
            "追溯: TRACE-REG-001", "追溯: TRACE-REG-002, TRACE-REG-001"
        )
        self._assert_markdown_error(reversed_text, "追溯引用顺序与追溯表不一致")

    def test_stray_headings_and_empty_nodes_rejected(self):
        stray = V2_MARKDOWN.replace(
            "  - [ ] 功能测试", "## 其他说明\n\n  - [ ] 功能测试"
        )
        self._assert_markdown_error(stray, "任务树中不得出现标题")
        empty_nodes = V2_MARKDOWN.replace(
            "\n## 需求追溯",
            "\n  - [ ] 回归测试\n\n- [ ] 空模块\n\n## 需求追溯",
        )
        errors = validate(empty_nodes)
        self.assertTrue(
            any("测试类型“回归测试”下没有任何用例" in error for error in errors), errors
        )
        self.assertTrue(
            any("功能模块“空模块”下没有任何用例" in error for error in errors), errors
        )

    def test_trace_section_is_not_scanned_for_list_markers(self):
        errors = validate(V2_MARKDOWN + "\n- 测试步骤\n  - 甲\n  - 乙\n")
        self.assertTrue(any("表外内容" in error for error in errors), errors)
        self.assertFalse(any("有序列表" in error for error in errors), errors)

    def test_markdown_cli_counts_parsed_cases(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cases.md"
            path.write_text(
                V2_MARKDOWN.replace(
                    "        - 提交注册。", "        - 提交注册。- [ ] 反向-"
                ),
                encoding="utf-8",
            )
            result = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "validate_markdown_lists.py"),
                    str(path),
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=subprocess_env(),
            )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("（1 条用例）", result.stdout)

    def test_excel_duplicate_trace_ids_rejected(self):
        def duplicate_row(workbook):
            sheet = workbook["需求追溯"]
            for column in range(1, 6):
                sheet.cell(row=4, column=column, value=sheet.cell(row=3, column=column).value)

        self._assert_excel_error(duplicate_row, "追溯 ID 重复")

        def pending_same_id(workbook):
            sheet = workbook.create_sheet("待确认项")
            sheet["A1"] = "待确认项清单"
            sheet.merge_cells("A1:F1")
            for index, header in enumerate(PENDING_HEADERS, start=1):
                sheet.cell(row=2, column=index, value=header)
            values = [
                "TRACE-REG-001",
                "注册需求.md",
                "1.2 超时",
                "未定义",
                "暂按失败处理",
                "是否重试？",
            ]
            for index, value in enumerate(values, start=1):
                sheet.cell(row=3, column=index, value=value)
            sheet.freeze_panes = "A3"
            sheet.auto_filter.ref = "A2:F3"

        self._assert_excel_error(pending_same_id, "追溯 ID 重复")

    def test_excel_titles_and_info_layout_required(self):
        self._assert_excel_error(
            lambda workbook: workbook["测试用例"].__setitem__("A1", None),
            "主工作表 A1 标题不能为空",
        )
        self._assert_excel_error(
            lambda workbook: workbook["生成说明"].__setitem__("A1", None),
            "生成说明 A1 标题不能为空",
        )
        self._assert_excel_error(
            lambda workbook: workbook["生成说明"].unmerge_cells("A1:B1"),
            "A1:B1 合并标题区域",
        )
        self._assert_excel_error(
            lambda workbook: workbook["测试用例"].__setitem__("L1", "=1+1"),
            "包含公式",
        )
        self._assert_excel_error(
            lambda workbook: workbook["需求追溯"].__setitem__("E2", "=1+1"),
            "包含公式",
        )

    def test_sync_accepts_stale_count_declaration(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            markdown = directory_path / "cases.md"
            excel = directory_path / "cases.xlsx"
            text = render_markdown(normalize_baseline(valid_ir_two_cases()))
            text = text.replace("> 用例数量: 用例 2 条", "> 用例数量: 用例 1 条")
            markdown.write_text(text, encoding="utf-8")
            save_workbook_atomic(build_workbook_from_ir(valid_ir_two_cases()), excel)
            code = sync_artifacts_main(
                ["--from-markdown", str(markdown), "--excel", str(excel)]
            )
            self.assertEqual(code, 0)
            self.assertIn(
                "> 用例数量: 用例 2 条", markdown.read_text(encoding="utf-8")
            )

    def test_sync_supports_custom_sheet_name(self):
        ir = valid_ir()
        ir["meta"]["sheet_name"] = "自定义用例"
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            markdown = directory_path / "cases.md"
            excel = directory_path / "cases.xlsx"
            save_workbook_atomic(build_workbook_from_ir(ir), excel)
            markdown.write_text(render_markdown(normalize_baseline(ir)), encoding="utf-8")
            code = sync_artifacts_main(
                [
                    "--from-markdown",
                    str(markdown),
                    "--excel",
                    str(excel),
                    "--sheet-name",
                    "自定义用例",
                ]
            )
            self.assertEqual(code, 0)
            self.assertEqual(load_workbook(excel).sheetnames[0], "自定义用例")

    def test_sync_drops_stale_design_method_warning(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            markdown = directory_path / "cases.md"
            excel = directory_path / "cases.xlsx"
            markdown.write_text(
                render_markdown(normalize_baseline(valid_ir_two_cases())).replace(
                    "[TC-ZC-002]", "[TC-NEW]"
                ),
                encoding="utf-8",
            )
            save_workbook_atomic(build_workbook_from_ir(valid_ir()), excel)
            _, errors, warnings = prepare_sync_ir(
                markdown,
                excel,
                assign_ids=True,
                id_prefix="TC-ZC",
                default_design_method="等价类划分",
            )
        self.assertEqual(errors, [])
        self.assertFalse(
            any("缺少 design_method" in warning for warning in warnings), warnings
        )

    def test_render_creates_missing_directories(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            baseline = directory_path / "ir.json"
            self._write_ir(baseline, valid_ir())
            markdown = directory_path / "TCD" / "2026-11-24" / "Markdown" / "cases.md"
            excel = directory_path / "TCD" / "2026-11-24" / "Excel" / "cases.xlsx"
            code = render_cases_main(
                [
                    "--baseline",
                    str(baseline),
                    "--markdown",
                    str(markdown),
                    "--excel",
                    str(excel),
                ]
            )
            self.assertEqual(code, 0)
            self.assertTrue(markdown.is_file())
            self.assertTrue(excel.is_file())

    def test_render_accepts_stale_traceability_links(self):
        ir = valid_ir()
        ir["traceability"][0]["case_ids"] = ["TC-ZC-999"]
        rendered, errors, _ = build_render_ir(ir)
        self.assertEqual(errors, [])
        self.assertEqual(rendered["traceability"][0]["case_ids"], ["TC-ZC-001"])

    def test_render_handles_bracketed_titles(self):
        with tempfile.TemporaryDirectory() as directory:
            code, _, _, excel = self._render(Path(directory), bracketed_ir())
            self.assertEqual(code, 0)
            self.assertEqual(validate_excel_file(excel), [])

    def test_merge_does_not_resurrect_deleted_fields(self):
        text = render_markdown(normalize_baseline(valid_ir()))
        block = "- 后置操作\n\n        - 删除本次注册的账号。"
        self.assertIn(block, text)
        text = text.replace(block, "")
        with tempfile.TemporaryDirectory() as directory:
            markdown = Path(directory) / "cases.md"
            excel = Path(directory) / "cases.xlsx"
            markdown.write_text(text, encoding="utf-8")
            save_workbook_atomic(build_workbook_from_ir(valid_ir()), excel)
            merged, errors, _ = merge_artifacts_to_ir(md_path=markdown, excel_path=excel)
        self.assertEqual(errors, [])
        self.assertEqual(merged["cases"][0]["post_operations"], [])

    def test_bundle_baseline_compares_and_reads_raw(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            code, baseline, markdown, excel = self._render(directory_path, valid_ir())
            self.assertEqual(code, 0)
            markdown.write_text(
                markdown.read_text(encoding="utf-8").replace(
                    "[合法手机号]", "[非法手机号]"
                ),
                encoding="utf-8",
            )
            errors = validate_bundle(
                markdown_path=markdown, excel_path=excel, baseline_path=baseline
            )
            self.assertTrue(
                any("基线 IR" in error and "title" in error for error in errors),
                errors,
            )

            raw = json.loads(baseline.read_text(encoding="utf-8"))
            raw["meta"]["case_count"] = 999
            baseline.write_text(json.dumps(raw, ensure_ascii=False), encoding="utf-8")
            errors = validate_bundle(markdown_path=markdown, baseline_path=baseline)
        self.assertTrue(any("case_count" in error for error in errors), errors)

    def test_corrupted_excel_reports_error(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "bad.xlsx"
            path.write_bytes(b"not a zip")
            ir, errors, _ = parse_excel_to_ir(path)
        self.assertEqual(ir, {})
        self.assertTrue(any("无法读取 Excel 文件" in error for error in errors), errors)

    def test_commit_reports_backup_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "cases.md"
            target.write_text("旧文件", encoding="utf-8")
            temp = Path(directory) / "cases.tmp"
            temp.write_text("新文件", encoding="utf-8")
            with mock.patch.object(
                validation_utils, "copy_to_temp", side_effect=OSError("占用")
            ):
                errors = commit_files_atomic([(target, temp)])
            self.assertTrue(any("备份文件失败" in error for error in errors), errors)
            self.assertEqual(target.read_text(encoding="utf-8"), "旧文件")


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

    def test_formula_like_input_remains_plain_text(self):
        ir = pending_ir()
        ir["traceability"][0].update(
            {
                "id": "=1+1",
                "source": "+1+1",
                "location": "-1+1",
                "rule": "+1+1",
            }
        )
        ir["cases"][0]["traceability_refs"] = ["=1+1"]
        ir["cases"][0].update(
            {
                "title": "-1+1",
                "pre_operations": ["@SUM(1)"],
                "test_data": ["=1+1"],
                "steps": ["+1+1"],
                "expected_results": ["-1+1"],
                "post_operations": ["@SUM(1)"],
            }
        )
        ir["meta"]["requirement_name"] = "=1+1"
        ir["meta"]["coverage"] = "=1+1"
        ir["meta"]["notes"] = "+1+1"
        ir["traceability"][1].update(
            {
                "source": "+1+1",
                "location": "-1+1",
                "issue": "=1+1",
                "assumption": "+1+1",
                "question": "-1+1",
            }
        )

        workbook = build_workbook_from_ir(ir)
        text_cells = {
            "测试用例": ("A3", "C3", "G3", "H3", "I3", "J3", "K3"),
            "生成说明": ("B2", "B5", "B6"),
            "需求追溯": ("A3", "B3", "C3", "D3", "E3"),
            "待确认项": ("A3", "B3", "C3", "D3", "E3", "F3"),
        }
        for sheet_name, cells in text_cells.items():
            for cell in cells:
                with self.subTest(sheet=sheet_name, cell=cell):
                    self.assertEqual(
                        workbook[sheet_name][cell].data_type,
                        "s",
                        f"{sheet_name}!{cell} 必须按纯文本保存",
                    )

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cases.xlsx"
            save_workbook_atomic(workbook, path)
            self.assertEqual(validate_excel_file(path), [])
            loaded = load_workbook(path)
            self.assertEqual(loaded["测试用例"]["C3"].value, "-1+1")
            self.assertEqual(loaded["测试用例"]["H3"].value, "=1+1")
            self.assertEqual(loaded["生成说明"]["B2"].value, "=1+1")
            self.assertEqual(loaded["生成说明"]["B5"].value, "=1+1")
            self.assertEqual(loaded["待确认项"]["F3"].value, "-1+1")

            workbook = build_workbook_from_ir(valid_ir())
            workbook["测试用例"]["G3"] = "=1+1"
            save_workbook_atomic(workbook, path)
            _, errors, _ = parse_excel_to_ir(path)
        self.assertTrue(any("包含公式" in error for error in errors), errors)

    def test_workbook_lock_is_friendly_and_cleans_temp(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            path = directory_path / "cases.xlsx"
            original_replace = os.replace

            def locked_replace(source, target):
                if Path(target) == path:
                    raise PermissionError(13, "locked", str(path))
                return original_replace(source, target)

            with mock.patch(
                "create_excel_cases.os.replace", side_effect=locked_replace
            ):
                with self.assertRaises(FileLockedError) as context:
                    save_workbook_atomic(build_workbook_from_ir(valid_ir()), path)
            self.assertIn("文件被占用", str(context.exception))
            self.assertIn(str(path), str(context.exception))
            self.assertEqual(list(directory_path.iterdir()), [])

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


class DraftSettlementTests(unittest.TestCase):
    """Scenario 1: draft traceability links settle from cases[].traceability_refs."""

    def _settled_ir(self, status: str, case_ids) -> dict:
        ir = valid_ir()
        ir["traceability"][0].update({"status": status, "case_ids": case_ids})
        return ir

    def test_render_cases_settles_draft_links_without_assign_ids(self):
        for status, case_ids in (
            ("已覆盖", []),
            ("待覆盖", ["TC-ZC-001"]),
            ("已覆盖", ["", "OLD-1", "OLD-1"]),
            ("待覆盖", ["", "OLD-1", "OLD-1"]),
        ):
            with self.subTest(status=status, case_ids=case_ids):
                with tempfile.TemporaryDirectory() as directory:
                    directory_path = Path(directory)
                    baseline = directory_path / "ir.json"
                    baseline.write_text(
                        json.dumps(self._settled_ir(status, case_ids), ensure_ascii=False),
                        encoding="utf-8",
                    )
                    markdown = directory_path / "cases.md"
                    excel = directory_path / "cases.xlsx"
                    code = render_cases_main(
                        [
                            "--baseline",
                            str(baseline),
                            "--markdown",
                            str(markdown),
                            "--excel",
                            str(excel),
                        ]
                    )
                    self.assertEqual(code, 0)
                    self.assertIn("TC-ZC-001", markdown.read_text(encoding="utf-8"))
                    self.assertEqual(validate_excel_file(excel), [])

    def test_single_format_clis_settle_draft_links(self):
        for status, case_ids in (
            ("已覆盖", []),
            ("待覆盖", ["", "OLD-1", "OLD-1"]),
        ):
            with self.subTest(status=status):
                with tempfile.TemporaryDirectory() as directory:
                    directory_path = Path(directory)
                    baseline = directory_path / "ir.json"
                    baseline.write_text(
                        json.dumps(self._settled_ir(status, case_ids), ensure_ascii=False),
                        encoding="utf-8",
                    )
                    markdown = directory_path / "cases.md"
                    excel = directory_path / "cases.xlsx"
                    self.assertEqual(
                        render_markdown_main(
                            ["--baseline", str(baseline), "--output", str(markdown)]
                        ),
                        0,
                    )
                    self.assertEqual(
                        create_excel_cases.main(
                            ["--baseline", str(baseline), "--output", str(excel)]
                        ),
                        0,
                    )
                    self.assertIn("TC-ZC-001", markdown.read_text(encoding="utf-8"))
                    self.assertEqual(validate_excel_file(excel), [])

    def test_prepare_render_ir_backfills_case_ids_from_refs(self):
        ir, errors, _ = prepare_render_ir(
            self._settled_ir("待覆盖", ["", "OLD-1", "OLD-1"])
        )
        self.assertEqual(errors, [])
        self.assertEqual(ir["traceability"][0]["case_ids"], ["TC-ZC-001"])
        self.assertEqual(ir["traceability"][0]["status"], "已覆盖")

    def test_pending_entry_with_case_ids_is_rejected_in_draft(self):
        ir = pending_ir()
        ir["traceability"][1]["case_ids"] = ["TC-ZC-001"]
        _, errors, _ = prepare_render_ir(ir)
        self.assertTrue(
            any("不得关联用例" in error for error in errors), errors
        )

    def test_non_array_case_ids_rejected_and_kept_out_of_result(self):
        entries, errors = validate_traceability_entries(
            [{"id": "T-1", "source": "s", "location": "l", "rule": "r",
              "status": "已覆盖", "case_ids": "TC-ZC-001"}],
            allow_unsettled_links=True,
        )
        self.assertEqual(entries, {})
        self.assertTrue(any("必须为数组" in error for error in errors), errors)

    def test_single_format_clis_still_reject_placeholder_ids(self):
        ir = valid_ir()
        ir["cases"][0]["id"] = "TC-NEW"
        ir["traceability"][0].update({"status": "待覆盖", "case_ids": []})
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            baseline = directory_path / "ir.json"
            baseline.write_text(json.dumps(ir, ensure_ascii=False), encoding="utf-8")
            markdown = directory_path / "cases.md"
            excel = directory_path / "cases.xlsx"
            self.assertEqual(
                render_markdown_main(
                    ["--baseline", str(baseline), "--output", str(markdown)]
                ),
                1,
            )
            self.assertEqual(
                create_excel_cases.main(
                    ["--baseline", str(baseline), "--output", str(excel)]
                ),
                1,
            )
            self.assertFalse(markdown.exists())
            self.assertFalse(excel.exists())

    def test_check_cli_still_blocks_unsettled_and_malformed_case_ids(self):
        cases = (
            ("已覆盖", []),
            ("待覆盖", ["TC-ZC-001"]),
            ("已覆盖", ["", "OLD-1"]),
            ("已覆盖", "TC-ZC-001"),
        )
        for status, case_ids in cases:
            with self.subTest(status=status, case_ids=case_ids):
                with tempfile.TemporaryDirectory() as directory:
                    baseline = Path(directory) / "ir.json"
                    baseline.write_text(
                        json.dumps(self._settled_ir(status, case_ids), ensure_ascii=False),
                        encoding="utf-8",
                    )
                    self.assertEqual(
                        case_baseline.main(["check", "--baseline", str(baseline)]), 1
                    )


class DraftCheckCliTests(unittest.TestCase):
    """`case_baseline.py check --draft` is a read-only draft pre-check."""

    def _write(self, directory: Path, ir: dict) -> Path:
        baseline = directory / "ir.json"
        baseline.write_text(json.dumps(ir, ensure_ascii=False), encoding="utf-8")
        return baseline

    def _draft_ir(self) -> dict:
        ir = valid_ir()
        ir["traceability"][0].update({"status": "待覆盖"})
        ir["traceability"][0].pop("case_ids")
        return ir

    def test_draft_passes_omitted_case_ids_and_reads_only(self):
        ir = self._draft_ir()
        ir["cases"][0]["group"] = ir["cases"][0]["module"]
        with tempfile.TemporaryDirectory() as directory:
            baseline = self._write(Path(directory), ir)
            digest = hashlib.sha256(baseline.read_bytes()).hexdigest()
            stdout = io.StringIO()
            stderr = io.StringIO()
            with redirect_stdout(stdout), redirect_stderr(stderr):
                code = case_baseline.main(
                    ["check", "--baseline", str(baseline), "--draft"]
                )
            self.assertEqual(code, 0)
            self.assertIn("校验通过（起草态预检）：共 1 条用例，已覆盖追溯 1 条，待确认 0 条。", stdout.getvalue())
            self.assertNotIn("模拟补号", stdout.getvalue())
            self.assertIn("提示：模块“注册模块”下存在同名分组“注册模块”", stderr.getvalue())
            self.assertEqual(hashlib.sha256(baseline.read_bytes()).hexdigest(), digest)

    def test_default_check_still_blocks_draft_ir(self):
        with tempfile.TemporaryDirectory() as directory:
            baseline = self._write(Path(directory), self._draft_ir())
            with redirect_stderr(io.StringIO()):
                self.assertEqual(case_baseline.main(["check", "--baseline", str(baseline)]), 1)

    def test_draft_assign_ids_reports_simulated_ids_without_writing(self):
        ir = valid_ir_two_cases()
        ir["cases"][1]["id"] = "TC-NEW"
        ir["traceability"][0]["status"] = "待覆盖"
        ir["traceability"][0].pop("case_ids")
        with tempfile.TemporaryDirectory() as directory:
            baseline = self._write(Path(directory), ir)
            digest = hashlib.sha256(baseline.read_bytes()).hexdigest()
            stdout = io.StringIO()
            with redirect_stdout(stdout):
                code = case_baseline.main(
                    ["check", "--baseline", str(baseline), "--draft", "--assign-ids"]
                )
            self.assertEqual(code, 0)
            self.assertIn(
                "模拟补号 1 条：TC-ZC-002（只读预检，未写回 IR）。", stdout.getvalue()
            )
            self.assertEqual(hashlib.sha256(baseline.read_bytes()).hexdigest(), digest)

    def test_draft_require_excel_fields_blocks_missing_design_method(self):
        ir = self._draft_ir()
        ir["cases"][0]["design_method"] = ""
        with tempfile.TemporaryDirectory() as directory:
            baseline = self._write(Path(directory), ir)
            stderr = io.StringIO()
            with redirect_stderr(stderr):
                code = case_baseline.main(
                    [
                        "check",
                        "--baseline",
                        str(baseline),
                        "--draft",
                        "--require-excel-fields",
                    ]
                )
            self.assertEqual(code, 1)
            self.assertIn("design_method", stderr.getvalue())

    def test_draft_blocks_unreferenced_trace_entry(self):
        ir = self._draft_ir()
        ir["traceability"].append(
            {
                "id": "TRACE-REG-002",
                "source": "注册需求.md",
                "location": "1.3 注册频率",
                "rule": "同一手机号每日最多注册 3 次",
                "status": "已覆盖",
            }
        )
        with tempfile.TemporaryDirectory() as directory:
            baseline = self._write(Path(directory), ir)
            stderr = io.StringIO()
            with redirect_stderr(stderr):
                code = case_baseline.main(
                    ["check", "--baseline", str(baseline), "--draft"]
                )
            self.assertEqual(code, 1)
            self.assertIn("TRACE-REG-002", stderr.getvalue())
            self.assertIn("待覆盖", stderr.getvalue())

    def test_assign_ids_without_draft_is_rejected_by_parser(self):
        with tempfile.TemporaryDirectory() as directory:
            baseline = self._write(Path(directory), valid_ir())
            with redirect_stderr(io.StringIO()):
                with self.assertRaises(SystemExit) as context:
                    case_baseline.main(
                        ["check", "--baseline", str(baseline), "--assign-ids"]
                    )
            self.assertEqual(context.exception.code, 2)


class StructureWarningTests(unittest.TestCase):
    """Scenario 8: module/group structure guards and non-blocking warnings."""

    def test_validate_rejects_case_type_as_module(self):
        for module in ("回归测试", "功能测试"):
            with self.subTest(module=module):
                ir = valid_ir()
                ir["cases"][0]["module"] = module
                errors = validate_baseline(ir)
                self.assertTrue(
                    any("不能使用测试类型名" in error for error in errors), errors
                )
                self.assertEqual(validate_baseline(ir, require_tree_fields=False), [])

    def test_collect_structure_warnings_dedupes_and_tolerates_bad_items(self):
        cases = [
            {"module": "模块甲", "group": "模块甲"},
            {"module": "模块甲", "group": "模块甲"},
            {"module": "模块甲", "group": "分组一"},
            {"module": "模块乙", "group": "模块乙"},
            "not-a-dict",
            None,
            {"module": None, "group": None},
        ]
        warnings = collect_structure_warnings(cases)
        self.assertEqual(len(warnings), 2)
        self.assertIn("模块“模块甲”下存在同名分组“模块甲”", warnings[0])
        self.assertIn("模块“模块乙”下存在同名分组“模块乙”", warnings[1])
        self.assertEqual(collect_structure_warnings("not-a-list"), [])

    def test_prepare_render_ir_warns_without_blocking(self):
        ir = valid_ir()
        ir["cases"][0]["group"] = "注册模块"
        settled, errors, warnings = prepare_render_ir(ir)
        self.assertEqual(errors, [])
        self.assertEqual(len(settled["cases"]), 1)
        self.assertTrue(any("同名分组" in warning for warning in warnings), warnings)

    def test_prepare_sync_ir_warns_without_blocking(self):
        ir = valid_ir()
        ir["cases"][0]["group"] = "注册模块"
        settled, errors, _ = prepare_render_ir(ir)
        self.assertEqual(errors, [])
        with tempfile.TemporaryDirectory() as directory:
            markdown = Path(directory) / "cases.md"
            markdown.write_text(render_markdown(settled), encoding="utf-8")
            _, sync_errors, warnings = prepare_sync_ir(markdown)
        self.assertEqual(sync_errors, [])
        self.assertTrue(any("同名分组" in warning for warning in warnings), warnings)


class ZeroLeakGuardTests(unittest.TestCase):
    """Scenario 2: an entry no case references can never render as covered."""

    def test_unreferenced_entry_with_stale_case_ids_blocks_render(self):
        for status in ("已覆盖", "待覆盖"):
            with self.subTest(status=status):
                ir = valid_ir()
                ir["traceability"].append(
                    {
                        "id": "TRACE-REG-002",
                        "source": "注册需求.md",
                        "location": "1.3 验证码",
                        "rule": "验证码 5 分钟内有效",
                        "status": status,
                        "case_ids": ["TC-ZC-999"],
                    }
                )
                with tempfile.TemporaryDirectory() as directory:
                    directory_path = Path(directory)
                    baseline = directory_path / "ir.json"
                    baseline.write_text(
                        json.dumps(ir, ensure_ascii=False), encoding="utf-8"
                    )
                    markdown = directory_path / "cases.md"
                    code = render_cases_main(
                        ["--baseline", str(baseline), "--markdown", str(markdown)]
                    )
                    self.assertEqual(code, 1)
                    self.assertFalse(markdown.exists())

    def test_prepare_render_ir_falls_back_to_pending_coverage(self):
        ir = valid_ir()
        ir["traceability"].append(
            {
                "id": "TRACE-REG-002",
                "source": "注册需求.md",
                "location": "1.3 验证码",
                "rule": "验证码 5 分钟内有效",
                "status": "已覆盖",
                "case_ids": ["TC-ZC-999"],
            }
        )
        _, errors, warnings = prepare_render_ir(ir)
        self.assertTrue(
            any("TRACE-REG-002" in error and "待覆盖" in error for error in errors),
            errors,
        )
        self.assertTrue(
            any("TRACE-REG-002" in warning and "回退" in warning for warning in warnings),
            warnings,
        )


class CascadeErrorTests(unittest.TestCase):
    """Scenario 3: one broken entry must not spray 'missing trace ID' errors."""

    def _draft_errors(self, mutate) -> list[str]:
        ir = valid_ir()
        mutate(ir)
        return validate_baseline(
            ir,
            allow_unsettled_coverage=True,
            require_finalized_traceability=False,
            check_traceability_links=False,
        )

    def test_covered_with_empty_case_ids_has_no_missing_id_error(self):
        ir = valid_ir()
        ir["traceability"][0]["case_ids"] = []
        errors = validate_baseline(ir)
        self.assertTrue(
            any("必须至少关联一个用例" in error for error in errors), errors
        )
        self.assertFalse(any("不存在" in error for error in errors), errors)

    def test_missing_rule_reports_only_the_entry_error(self):
        errors = self._draft_errors(lambda ir: ir["traceability"][0].pop("rule"))
        self.assertTrue(any("rule 不能为空" in error for error in errors), errors)
        self.assertFalse(any("不存在" in error for error in errors), errors)

    def test_missing_source_reports_only_the_entry_error(self):
        errors = self._draft_errors(lambda ir: ir["traceability"][0].update(source=""))
        self.assertTrue(any("source 与 location 不能为空" in error for error in errors), errors)
        self.assertFalse(any("不存在" in error for error in errors), errors)

    def test_referenced_pending_entry_still_reports_pending_reference(self):
        ir = pending_ir()
        ir["traceability"][1]["case_ids"] = ["TC-ZC-001"]
        ir["cases"][0]["traceability_refs"] = ["TRACE-REG-001", "TRACE-REG-Q01"]
        errors = validate_baseline(ir)
        self.assertTrue(
            any("不得引用待确认条目 'TRACE-REG-Q01'" in error for error in errors),
            errors,
        )
        self.assertFalse(any("不存在" in error for error in errors), errors)

    def test_typo_id_still_reports_missing_trace_id(self):
        errors = self._draft_errors(
            lambda ir: ir["cases"][0].update(traceability_refs=["TRACE-REG-999"])
        )
        self.assertTrue(
            any("引用了不存在的追溯 ID 'TRACE-REG-999'" in error for error in errors),
            errors,
        )


class ExampleAssetTests(unittest.TestCase):
    """Scenario 4: the shipped sample IR must stay checkable and renderable."""

    def test_sample_passes_check_and_renders_four_sheets(self):
        sample = ROOT / "assets" / "example-cases-baseline.json"
        before = sample.read_bytes()
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            baseline = directory_path / "ir.json"
            baseline.write_bytes(before)
            self.assertEqual(
                case_baseline.main(["check", "--baseline", str(baseline)]), 0
            )
            markdown = directory_path / "cases.md"
            excel = directory_path / "cases.xlsx"
            self.assertEqual(
                render_cases_main(
                    [
                        "--baseline",
                        str(baseline),
                        "--markdown",
                        str(markdown),
                        "--excel",
                        str(excel),
                    ]
                ),
                0,
            )
            self.assertEqual(validate_excel_file(excel), [])
            self.assertEqual(
                load_workbook(excel).sheetnames,
                ["测试用例", "生成说明", "需求追溯", "待确认项"],
            )
            self.assertEqual(
                validate_bundle(markdown_path=markdown, excel_path=excel), []
            )
        self.assertEqual(sample.read_bytes(), before)


class WritePreflightTests(unittest.TestCase):
    """Scenario 5: permission, directory and lock pre-flight attribution."""

    def test_missing_target_with_unwritable_parent(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "cases.md"
            with mock.patch.object(
                validation_utils, "_is_parent_writable", return_value=False
            ):
                errors = _check_targets_writable([target])
        self.assertEqual(len(errors), 1)
        self.assertIn("输出目录不可写", errors[0])
        self.assertNotIn("被占用", errors[0])

    def test_replace_reports_unwritable_parent_only_when_unwritable(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "cases.md"
            temp = Path(directory) / "cases.tmp"
            temp.write_text("新", encoding="utf-8")
            with mock.patch.object(
                validation_utils, "_is_parent_writable", return_value=False
            ), mock.patch.object(
                validation_utils.os,
                "replace",
                side_effect=PermissionError(13, "denied", str(target)),
            ):
                with self.assertRaises(PermissionError) as context:
                    replace_file_atomic(temp, target)
            self.assertIn("输出目录不可写", str(context.exception))

            temp.write_text("新", encoding="utf-8")
            with mock.patch.object(
                validation_utils, "_is_parent_writable", return_value=True
            ), mock.patch.object(
                validation_utils.os,
                "replace",
                side_effect=PermissionError(13, "denied", str(target)),
            ):
                with self.assertRaises(PermissionError) as context:
                    replace_file_atomic(temp, target)
            self.assertNotIn("输出目录不可写", str(context.exception))
            self.assertNotIn("被占用", str(context.exception))

    def test_readonly_target_skips_backup_and_replace(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "cases.md"
            target.write_text("旧", encoding="utf-8")
            temp = sibling_temp_path(target)
            temp.write_text("新", encoding="utf-8")
            with mock.patch.object(
                validation_utils, "_is_readonly_or_unwritable", return_value=True
            ), mock.patch.object(
                validation_utils, "copy_to_temp"
            ) as copy_mock, mock.patch.object(
                validation_utils, "replace_file_atomic"
            ) as replace_mock:
                errors = commit_files_atomic([(target, temp)])
        self.assertEqual(len(errors), 1)
        self.assertIn("文件不可写", errors[0])
        self.assertNotIn("被占用", errors[0])
        copy_mock.assert_not_called()
        replace_mock.assert_not_called()

    def test_non_permission_oserror_keeps_its_own_text(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "cases.md"
            target.write_text("x", encoding="utf-8")
            with mock.patch.object(
                validation_utils.os, "name", "nt"
            ), mock.patch.object(
                validation_utils, "_is_readonly_or_unwritable", return_value=False
            ), mock.patch.object(
                Path, "open", side_effect=OSError(28, "No space left on device")
            ):
                errors = _check_targets_writable([target])
        self.assertEqual(len(errors), 1)
        self.assertNotIn("被占用", errors[0])
        self.assertIn("No space left", errors[0])

    def test_commit_passes_through_readonly_and_dir_permission_errors(self):
        for message, expected in (
            ("文件不可写（只读或权限不足）：x", "文件不可写"),
            ("输出目录不可写（权限不足），无法写入：x", "输出目录不可写"),
        ):
            with self.subTest(expected=expected):
                with tempfile.TemporaryDirectory() as directory:
                    target = Path(directory) / "cases.md"
                    target.write_text("旧", encoding="utf-8")
                    temp = sibling_temp_path(target)
                    temp.write_text("新", encoding="utf-8")
                    with mock.patch.object(
                        validation_utils,
                        "replace_file_atomic",
                        side_effect=PermissionError(message),
                    ):
                        errors = commit_files_atomic([(target, temp)])
                self.assertEqual(len(errors), 1)
                self.assertIn(expected, errors[0])
                self.assertNotIn("文件被占用", errors[0])

    def test_windows_share_lock_on_second_target_touches_nothing(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            first = directory_path / "a.md"
            second = directory_path / "b.md"
            first.write_text("旧甲", encoding="utf-8")
            second.write_text("旧乙", encoding="utf-8")
            first_temp = sibling_temp_path(first)
            first_temp.write_text("新甲", encoding="utf-8")
            second_temp = sibling_temp_path(second)
            second_temp.write_text("新乙", encoding="utf-8")
            first_bytes = first.read_bytes()

            original_open = Path.open
            share_violation = PermissionError(13, "locked", str(second))
            share_violation.winerror = 32

            def locked_open(self, mode="r", *args, **kwargs):
                if self == second and mode == "r+b":
                    raise share_violation
                return original_open(self, mode, *args, **kwargs)

            with mock.patch.object(
                validation_utils.os, "name", "nt"
            ), mock.patch.object(Path, "open", locked_open), mock.patch.object(
                validation_utils, "copy_to_temp"
            ) as copy_mock, mock.patch.object(
                validation_utils, "replace_file_atomic"
            ) as replace_mock:
                errors = commit_files_atomic(
                    [(first, first_temp), (second, second_temp)]
                )
            self.assertEqual(len(errors), 1)
            self.assertIn("文件被占用", errors[0])
            self.assertEqual(first.read_bytes(), first_bytes)
            copy_mock.assert_not_called()
            replace_mock.assert_not_called()

    def test_workbook_save_reports_readonly_target(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cases.xlsx"
            workbook = build_workbook_from_ir(valid_ir())
            with mock.patch.object(
                validation_utils, "_is_readonly_or_unwritable", return_value=True
            ), mock.patch.object(
                create_excel_cases.os,
                "replace",
                side_effect=PermissionError(13, "denied", str(path)),
            ):
                path.write_bytes(b"old")
                with self.assertRaises(PermissionError) as context:
                    save_workbook_atomic(workbook, path)
            self.assertIn("文件不可写", str(context.exception))
            self.assertEqual(path.read_bytes(), b"old")
            leftovers = [p.name for p in Path(directory).iterdir() if p.name != path.name]
            self.assertEqual(leftovers, [])


class RenderFailureHintTests(unittest.TestCase):
    """Scenario 5 tail: rollback-failure hints and temporary IR retention."""

    def _render_failure(self, rollback_fails: bool):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            baseline = directory_path / "ir.json"
            baseline.write_text(json.dumps(valid_ir(), ensure_ascii=False), encoding="utf-8")
            markdown = directory_path / "cases.md"
            excel = directory_path / "cases.xlsx"
            excel.write_bytes(b"old")
            if not rollback_fails:
                # A pre-existing first target rolls back by restoring its backup,
                # so only the never-existed case reaches the unlink-based rollback.
                markdown.write_text("旧 Markdown", encoding="utf-8")

            real_replace = validation_utils.replace_file_atomic

            def flaky_replace(source, target):
                # Staging also goes through replace_file_atomic, so fail on the
                # commit of the Excel target instead of counting calls.
                if Path(target) == excel:
                    raise PermissionError(13, "locked", str(target))
                return real_replace(source, target)

            argv = [
                "--baseline",
                str(baseline),
                "--markdown",
                str(markdown),
                "--excel",
                str(excel),
                "--force",
                # Without this a deleted IR would look like a pass; the failed
                # commit must skip cleanup and leave the IR in place.
                "--cleanup-baseline",
            ]
            stderr = io.StringIO()
            with mock.patch.object(
                validation_utils, "replace_file_atomic", side_effect=flaky_replace
            ):
                if rollback_fails:
                    original_unlink = Path.unlink

                    def failing_unlink(self, missing_ok=False):
                        if self.name == "cases.md":
                            raise OSError("回滚失败模拟")
                        return original_unlink(self, missing_ok=missing_ok)

                    with mock.patch.object(
                        Path, "unlink", failing_unlink
                    ), redirect_stderr(stderr):
                        code = render_cases_main(argv)
                else:
                    with redirect_stderr(stderr):
                        code = render_cases_main(argv)
            return code, stderr.getvalue(), baseline.exists()

    def test_failure_without_rollback_error_keeps_temp_ir(self):
        code, stderr, baseline_exists = self._render_failure(rollback_fails=False)
        self.assertEqual(code, 1)
        self.assertIn("其余目标文件均未变更", stderr)
        self.assertIn("临时 IR 已保留", stderr)
        self.assertTrue(baseline_exists)

    def test_failure_with_rollback_error_reports_manual_restore(self):
        code, stderr, baseline_exists = self._render_failure(rollback_fails=True)
        self.assertEqual(code, 1)
        self.assertIn("回滚失败", stderr)
        self.assertIn("临时 IR 仍保留", stderr)
        self.assertTrue(baseline_exists)

    def _sync_failure(self, rollback_fails: bool):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            baseline = directory_path / "ir.json"
            baseline.write_text(json.dumps(valid_ir(), ensure_ascii=False), encoding="utf-8")
            markdown = directory_path / "cases.md"
            excel = directory_path / "cases.xlsx"
            self.assertEqual(
                render_cases_main(
                    [
                        "--baseline",
                        str(baseline),
                        "--markdown",
                        str(markdown),
                        "--excel",
                        str(excel),
                    ]
                ),
                0,
            )
            markdown.write_text(
                markdown.read_text(encoding="utf-8").replace(
                    "[合法手机号]", "[非法手机号]"
                ),
                encoding="utf-8",
            )
            real_replace = validation_utils.replace_file_atomic

            def flaky_replace(source, target):
                if Path(target) == excel:
                    raise PermissionError(13, "locked", str(target))
                return real_replace(source, target)

            stderr = io.StringIO()
            with mock.patch.object(
                validation_utils, "replace_file_atomic", side_effect=flaky_replace
            ):
                if rollback_fails:
                    # Both targets exist, so rollback restores backups through
                    # os.replace; fail only that second replace of the Markdown.
                    real_os_replace = os.replace
                    seen = {"count": 0}

                    def failing_replace(source, target):
                        if Path(target) == markdown:
                            seen["count"] += 1
                            if seen["count"] == 2:
                                raise OSError("回滚失败模拟")
                        return real_os_replace(source, target)

                    with mock.patch.object(
                        validation_utils.os, "replace", failing_replace
                    ), redirect_stderr(stderr):
                        code = sync_artifacts_main(
                            ["--from-markdown", str(markdown), "--excel", str(excel)]
                        )
                else:
                    with redirect_stderr(stderr):
                        code = sync_artifacts_main(
                            ["--from-markdown", str(markdown), "--excel", str(excel)]
                        )
            return code, stderr.getvalue()

    def test_sync_failure_without_rollback_error_keeps_original_files(self):
        code, stderr = self._sync_failure(rollback_fails=False)
        self.assertEqual(code, 1)
        self.assertIn("其余目标文件均未变更", stderr)
        self.assertNotIn("回滚失败", stderr)

    def test_sync_failure_with_rollback_error_reports_manual_restore(self):
        code, stderr = self._sync_failure(rollback_fails=True)
        self.assertEqual(code, 1)
        self.assertIn("回滚失败", stderr)
        self.assertIn("请先按上方备份路径手工恢复", stderr)


class InspectSummaryTests(unittest.TestCase):
    """Scenario 6: the read-only inspect summary."""

    def _render(self, directory: Path, ir: dict) -> tuple[Path, Path]:
        baseline = directory / "ir.json"
        baseline.write_text(json.dumps(ir, ensure_ascii=False), encoding="utf-8")
        markdown = directory / "cases.md"
        excel = directory / "cases.xlsx"
        self.assertEqual(
            render_cases_main(
                [
                    "--baseline",
                    str(baseline),
                    "--markdown",
                    str(markdown),
                    "--excel",
                    str(excel),
                ]
            ),
            0,
        )
        return markdown, excel

    def test_single_sources_and_multi_source_sections(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            markdown, excel = self._render(directory_path, pending_ir())
            baseline = directory_path / "ir.json"
            before = {
                path: hashlib.sha256(path.read_bytes()).hexdigest()
                for path in (markdown, excel, baseline)
            }

            for kwargs, expected in (
                ({"markdown_path": markdown}, "=== Markdown ==="),
                ({"excel_path": excel}, "=== Excel ==="),
                ({"baseline_path": baseline}, "=== Baseline IR ==="),
            ):
                with self.subTest(source=expected):
                    lines, errors = inspect_artifacts(**kwargs)
                    self.assertEqual(errors, [])
                    self.assertEqual(lines[0], expected)

            lines, errors = inspect_artifacts(
                markdown_path=markdown, excel_path=excel, baseline_path=baseline
            )
            self.assertEqual(errors, [])
            self.assertEqual(
                [line for line in lines if line.startswith("=== ")],
                ["=== Baseline IR ===", "=== Markdown ===", "=== Excel ==="],
            )
            self.assertIn("工作表：测试用例、生成说明、需求追溯、待确认项", lines)
            for path, digest in before.items():
                self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(), digest)

    def test_requires_at_least_one_source(self):
        self.assertEqual(
            inspect_artifacts(), ([], ["必须提供 --markdown、--excel 或 --baseline 至少一个。"])
        )

    def _id_line_for(self, count: int) -> str:
        ir = valid_ir()
        base = ir["cases"][0]
        cases = []
        for index in range(1, count + 1):
            case = deepcopy(base)
            case.update({"id": f"TC-ZC-{index:03d}", "title": f"场景 {index}"})
            cases.append(case)
        ir["cases"] = cases
        ir["traceability"][0]["case_ids"] = [case["id"] for case in cases]
        with tempfile.TemporaryDirectory() as directory:
            baseline = Path(directory) / "ir.json"
            baseline.write_text(json.dumps(ir, ensure_ascii=False), encoding="utf-8")
            lines, errors = inspect_artifacts(baseline_path=baseline)
        self.assertEqual(errors, [])
        return next(line for line in lines if line.startswith("用例 ID："))

    def test_thirty_ids_are_listed_in_full(self):
        id_line = self._id_line_for(30)
        self.assertNotIn("…", id_line)
        self.assertIn("TC-ZC-001、TC-ZC-002", id_line)
        self.assertIn("TC-ZC-029、TC-ZC-030", id_line)
        self.assertIn("TC-ZC-015", id_line)
        self.assertTrue(id_line.endswith("TC-ZC-030"))

    def test_id_list_truncates_beyond_thirty(self):
        id_line = self._id_line_for(31)
        self.assertIn("TC-ZC-001、TC-ZC-002、TC-ZC-003、TC-ZC-004、TC-ZC-005", id_line)
        self.assertIn("TC-ZC-027、TC-ZC-028、TC-ZC-029、TC-ZC-030、TC-ZC-031", id_line)
        self.assertIn("（共 31 条）", id_line)
        self.assertNotIn("TC-ZC-015", id_line)

    def test_ids_expand_hits_and_report_misses(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            markdown, _ = self._render(directory_path, valid_ir())
            lines, errors = inspect_artifacts(
                markdown_path=markdown, ids=["TC-ZC-001", "TC-MISSING"]
            )
        self.assertEqual(errors, [])
        self.assertIn("--- TC-ZC-001 ---", lines)
        self.assertIn("未找到用例: TC-MISSING", lines)

    def test_ids_detail_includes_pre_operations_and_test_data(self):
        with tempfile.TemporaryDirectory() as directory:
            markdown, _ = self._render(Path(directory), valid_ir())
            lines, errors = inspect_artifacts(markdown_path=markdown, ids=["TC-ZC-001"])
        self.assertEqual(errors, [])
        positions = [
            next(index for index, line in enumerate(lines) if line.startswith(label + "："))
            for label in ("前置操作", "测试数据", "测试步骤", "预期结果", "后置操作")
        ]
        self.assertEqual(positions, sorted(positions))
        self.assertIn("前置操作：用户未注册。", lines)
        self.assertIn("测试数据：手机号: <测试手机号>", lines)

    def test_markdown_reports_no_coverage_field(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            baseline = directory_path / "ir.json"
            baseline.write_text(json.dumps(valid_ir(), ensure_ascii=False), encoding="utf-8")
            markdown = directory_path / "cases.md"
            self.assertEqual(
                render_cases_main(
                    ["--baseline", str(baseline), "--markdown", str(markdown)]
                ),
                0,
            )
            lines, errors = inspect_artifacts(markdown_path=markdown)
        self.assertEqual(errors, [])
        self.assertIn(
            "覆盖范围：Markdown 不含覆盖范围字段（以 Excel 生成说明为准）", lines
        )
        self.assertNotIn("覆盖范围：（空）", lines)

    def test_trace_ids_expand_covered_pending_and_misses(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            markdown, excel = self._render(directory_path, pending_ir())
            baseline = directory_path / "ir.json"
            before = {
                path: hashlib.sha256(path.read_bytes()).hexdigest()
                for path in (markdown, excel, baseline)
            }
            lines, errors = inspect_artifacts(
                markdown_path=markdown,
                excel_path=excel,
                baseline_path=baseline,
                trace_ids=["TRACE-REG-001", "TRACE-REG-Q01", "TRACE-MISSING"],
            )
            after = {
                path: hashlib.sha256(path.read_bytes()).hexdigest()
                for path in (markdown, excel, baseline)
            }
        self.assertEqual(errors, [])
        self.assertEqual(before, after)
        expected = [
            "--- TRACE-REG-001 ---",
            "状态：已覆盖",
            "来源：注册需求.md",
            "定位：1.1 手机号",
            "规则：手机号必须为11位数字",
            "关联用例：TC-ZC-001",
            "--- TRACE-REG-Q01 ---",
            "状态：待确认",
            "来源：注册需求.md",
            "定位：1.2 超时",
            "未明确项：超时行为未定义",
            "当前测试假设：暂按失败处理",
            "确认问题：超时是否重试？",
            "未找到追溯条目: TRACE-MISSING",
        ]
        for source in ("Markdown", "Excel", "Baseline IR"):
            with self.subTest(source=source):
                start = lines.index(f"=== {source} ===")
                end = next(
                    (
                        index
                        for index in range(start + 1, len(lines))
                        if lines[index].startswith("=== ")
                    ),
                    len(lines),
                )
                block = lines[start:end]
                position = -1
                for line in expected:
                    position = block.index(line, position + 1)
                pending_start = block.index("--- TRACE-REG-Q01 ---")
                pending_end = next(
                    (
                        index
                        for index in range(pending_start + 1, len(block))
                        if block[index].startswith(("--- ", "未找到追溯条目"))
                    ),
                    len(block),
                )
                pending_block = block[pending_start:pending_end]
                self.assertFalse(
                    any(
                        line.startswith(("规则：", "关联用例："))
                        for line in pending_block
                    ),
                    pending_block,
                )

    def test_excel_inspect_reports_missing_module_column(self):
        ir = valid_ir()
        ir["cases"][0]["title"] = "合法手机号注册成功"
        with tempfile.TemporaryDirectory() as directory:
            _, excel = self._render(Path(directory), ir)
            lines, errors = inspect_artifacts(excel_path=excel)
        self.assertEqual(errors, [])
        self.assertIn("按模块：Excel 不含模块列（未统计）", lines)
        self.assertFalse(any("（未标注）" in line for line in lines))

    def test_malformed_refs_do_not_crash(self):
        for refs in ("TRACE-REG-001", 123, None):
            with self.subTest(refs=refs):
                ir = valid_ir()
                ir["cases"][0]["traceability_refs"] = refs
                lines = case_baseline._inspect_summary(ir, "IR", ["TC-ZC-001"])
                self.assertIn("追溯引用：（无）", lines)

    def test_missing_excel_reports_error(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "missing.xlsx"
            lines, errors = inspect_artifacts(excel_path=path)
        self.assertEqual(lines, [])
        self.assertEqual(len(errors), 1)
        self.assertIn("Excel:", errors[0])
        self.assertIn("无法读取", errors[0])


class OpenpyxlBlockedInspectTests(unittest.TestCase):
    """Scenario 7: inspect must work without openpyxl except for --excel."""

    def test_inspect_markdown_and_baseline_without_openpyxl(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            blocker = directory_path / "blocker"
            blocker.mkdir()
            (blocker / "sitecustomize.py").write_text(
                "import sys\n"
                "\n"
                "class _BlockOpenpyxl:\n"
                "    def find_spec(self, fullname, path=None, target=None):\n"
                "        if fullname == 'openpyxl' or fullname.startswith('openpyxl.'):\n"
                "            raise ImportError('openpyxl blocked for inspect test')\n"
                "        return None\n"
                "\n"
                "sys.meta_path.insert(0, _BlockOpenpyxl())\n",
                encoding="utf-8",
            )
            baseline = directory_path / "ir.json"
            baseline.write_text(json.dumps(valid_ir(), ensure_ascii=False), encoding="utf-8")
            markdown = directory_path / "cases.md"
            self.assertEqual(
                render_cases_main(["--baseline", str(baseline), "--markdown", str(markdown)]),
                0,
            )
            env = {**subprocess_env(), "PYTHONPATH": str(blocker)}

            def run(*args: str) -> subprocess.CompletedProcess:
                return subprocess.run(
                    [sys.executable, *args],
                    cwd=directory_path,
                    env=env,
                    text=True,
                    encoding="utf-8",
                    capture_output=True,
                )

            self.assertNotEqual(run("-c", "import openpyxl").returncode, 0)

            script = str(ROOT / "scripts" / "case_baseline.py")
            for command, expected in (
                ((script, "inspect", "--markdown", str(markdown)), "=== Markdown ==="),
                ((script, "inspect", "--baseline", str(baseline)), "=== Baseline IR ==="),
                (
                    (
                        script,
                        "inspect",
                        "--markdown",
                        str(markdown),
                        "--trace-ids",
                        "TRACE-REG-001",
                    ),
                    "--- TRACE-REG-001 ---",
                ),
            ):
                with self.subTest(expected=expected):
                    result = run(*command)
                    self.assertEqual(result.returncode, 0, result.stderr)
                    self.assertIn(expected, result.stdout)
                    self.assertNotIn("openpyxl", result.stderr)

            excel = directory_path / "cases.xlsx"
            result = run(script, "inspect", "--excel", str(excel))
            self.assertEqual(result.returncode, 1)
            self.assertIn("openpyxl", result.stderr)


class DeadFunctionCleanupTests(unittest.TestCase):
    """Old-format support stays removed, and old input is rejected explicitly."""

    LEGACY_HEADERS = (
        "需求/追踪 ID",
        "用例 ID",
        "用例标题",
        "用例类型",
        "设计方法",
        "优先级",
        "前置条件",
        "测试数据",
        "执行步骤",
        "预期结果",
        "备注",
    )

    def test_legacy_helpers_are_gone(self):
        import traceability_matrix

        for name in (
            "load_traceability_matrix",
            "traceability_text",
            "LEGACY_TRACEABILITY_PATTERN",
            "TRACEABILITY_PATTERN",
            "REQUIRED_FIELDS",
            "configure_utf8_stderr",
            "entries_for_case_id",
            "entries_for_case_title",
            "final_matrix_errors",
        ):
            with self.subTest(module="traceability_matrix", name=name):
                self.assertFalse(hasattr(traceability_matrix, name))
        for name in ("V1_HEADERS", "_parse_v1_case_sheet", "_legacy_entries"):
            with self.subTest(module="case_baseline", name=name):
                self.assertFalse(hasattr(case_baseline, name))

    def test_legacy_workbook_headers_are_rejected(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "测试用例"
        sheet["A1"] = "测试用例 - 未指定日期 - [注册]"
        for index, header in enumerate(self.LEGACY_HEADERS, start=1):
            sheet.cell(row=2, column=index, value=header)
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "legacy.xlsx"
            workbook.save(target)
            _, errors, _ = parse_excel_to_ir(target)
        self.assertTrue(any("无法识别的工作表表头" in error for error in errors), errors)

    def test_alias_case_fields_are_rejected(self):
        for field in ("preconditions", "post_cleanup"):
            with self.subTest(field=field):
                ir = valid_ir()
                ir["cases"][0][field] = ["旧值"]
                errors = validate_baseline(ir)
                self.assertTrue(
                    any(field in error and "不支持字段" in error for error in errors),
                    errors,
                )

    def test_generated_note_alias_is_ignored(self):
        workbook = Workbook()
        create_excel_cases.build_info_sheet(workbook, {"generated_note": "旧备注"})
        self.assertFalse(workbook["生成说明"]["B6"].value)


if __name__ == "__main__":
    unittest.main()
