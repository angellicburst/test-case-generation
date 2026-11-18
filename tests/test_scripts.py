from __future__ import annotations

from copy import deepcopy
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from create_excel_cases import build_workbook  # noqa: E402
from traceability_matrix import load_traceability_matrix  # noqa: E402
from sync_traceability import sync_excel, sync_markdown  # noqa: E402
from validate_excel_cases import validate_excel_file  # noqa: E402
from validate_markdown_lists import fix, validate  # noqa: E402
from validate_traceability_bundle import validate_bundle  # noqa: E402


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


class MarkdownValidationTests(unittest.TestCase):
    def test_valid_document(self):
        self.assertEqual(validate(VALID_MARKDOWN, TRACEABILITY_MATRIX), [])

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
                build_workbook("测试用例", cases, loaded)
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
            build_workbook("测试用例", [excel_case()], TRACEABILITY_MATRIX).save(target)
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
            build_workbook("测试用例", cases, matrix).save(target)
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
            build_workbook("测试用例", cases, matrix_for_cases(*cases)).save(target)
            self.assertEqual(
                load_workbook(target)["测试用例"]["K3"].value,
                "[需求追溯] 来源：注册需求.md｜定位：2.1 手机号规则｜规则：手机号必须为11位数字\n补充备注",
            )

    def test_excel_supports_multiple_traceability_entries(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "multiple-traceability.xlsx"
            case = excel_case(traceability_refs=["TRACE-REG-002", "TRACE-REG-001"])
            matrix = multi_traceability_matrix()
            build_workbook("测试用例", [case], matrix).save(target)
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
            build_workbook("测试用例", cases, matrix).save(target)
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
            build_workbook("测试用例", [case], matrix).save(target)
            self.assertEqual(validate_excel_file(target, matrix), [])
            self.assertEqual(load_workbook(target)["测试用例"]["A3"].value, "US-1024")

    def test_excel_traceability_requires_exact_label(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "legacy-label.xlsx"
            cases = [excel_case()]
            matrix = matrix_for_cases(*cases)
            build_workbook("测试用例", cases, matrix).save(target)
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
            build_workbook("测试用例", cases, matrix, sheet_name="自定义").save(target)
            self.assertTrue(validate_excel_file(target, matrix))
            self.assertEqual(validate_excel_file(target, matrix, sheet_name="自定义"), [])

    def test_formula_like_input_remains_plain_text(self):
        markers = ["=1+1", "+1+1", "-1+1", "@SUM(1)"]
        case = excel_case(requirement_id=markers[0], id=markers[1], title=markers[2], type=markers[3])
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "text.xlsx"
            build_workbook(
                "=标题",
                [case],
                matrix_for_cases(case),
                meta={"requirement_name": "=需求", "notes": "@说明"},
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
            build_workbook("测试用例", cases, matrix).save(target)
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
            build_workbook("测试用例", cases, matrix).save(target)
            workbook = load_workbook(target)
            sheet = workbook["测试用例"]
            sheet.freeze_panes = None
            sheet.auto_filter.ref = "A2:J3"
            workbook.save(target)
            errors = validate_excel_file(target, matrix)
            self.assertTrue(any("冻结窗格至 A3" in error for error in errors))
            self.assertTrue(any("筛选范围必须为 A2:K3" in error for error in errors))

    def test_excel_required_fields_duplicates_and_priority(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "invalid.xlsx"
            cases = [
                excel_case(priority="P5", preconditions=""),
                excel_case(title="提交合法注册"),
            ]
            matrix = matrix_for_cases(*cases)
            build_workbook("测试用例", cases, matrix).save(target)
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

    def test_excel_type_design_method_and_stale_matrix_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "invalid-enum.xlsx"
            case = excel_case(type="异常测试", design_method="判定表")
            matrix = matrix_for_cases(case)
            matrix["TRACE-REG-001"]["case_ids"].append("TC-REG-STALE")
            build_workbook("测试用例", [case], matrix).save(target)
            errors = validate_excel_file(target, matrix)
            self.assertTrue(any("用例类型无效" in error for error in errors))
            self.assertTrue(any("设计方法无效" in error for error in errors))
            self.assertTrue(any("TC-REG-STALE”未出现在工作表中" in error for error in errors))

    def test_excel_optional_post_cleanup_stays_in_notes(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "cleanup.xlsx"
            case = excel_case(notes="后置清理：删除测试账号并恢复配置。")
            matrix = matrix_for_cases(case)
            build_workbook("测试用例", [case], matrix).save(target)
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
                json.dumps({"meta": {"sheet_name": "自定义"}, "cases": cases}, ensure_ascii=False),
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
            build_workbook("测试用例", [case], matrix).save(excel)
            matrix["TRACE-REG-001"]["rule"] = "手机号必须为11位且不得包含空格"
            write_matrix(matrix_path, matrix)

            self.assertTrue(validate_bundle(matrix_path, markdown, excel))
            self.assertEqual(sync_markdown(markdown, matrix), [])
            self.assertEqual(sync_excel(excel, matrix, "测试用例"), [])
            self.assertEqual(validate_bundle(matrix_path, markdown, excel), [])

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
            build_workbook("测试用例", [excel_case()], matrix).save(excel)
            matrix["TRACE-REG-001"]["case_titles"] = []
            write_matrix(matrix_path, matrix)
            errors = validate_bundle(matrix_path, markdown, excel)
            self.assertTrue(any("缺少 Markdown 用例关联" in error for error in errors))

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
            build_workbook("测试用例", [excel_case()], matrix).save(excel)
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
            build_workbook("测试用例", [excel_case()], matrix).save(excel)
            write_matrix(matrix_path, matrix)

            errors = validate_bundle(
                matrix_path,
                markdown,
                excel,
                coverage_only=True,
            )
            self.assertTrue(any("缺少 Markdown 用例关联" in error for error in errors))
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


if __name__ == "__main__":
    unittest.main()
