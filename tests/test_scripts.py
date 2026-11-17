from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from create_excel_cases import build_workbook  # noqa: E402
from validate_excel_cases import validate_excel_file  # noqa: E402
from validate_markdown_lists import fix, validate  # noqa: E402


VALID_MARKDOWN = """# 测试用例 - 2026-11-17 - [注册]
> 用例数量: 用例 1 条

- [ ] 注册模块
  - [ ] 功能测试
    - [ ] 正向-提交合法注册-P0
      - 前置条件
        - 用户未注册。
      - 测试步骤
        - 提交注册。
      - 预期结果
        - 注册成功。
"""


def excel_case(**overrides):
    case = {
        "requirement_id": "REQ-1",
        "id": "TC-REG-001",
        "title": "提交合法注册",
        "type": "功能测试",
        "design_method": "等价类划分",
        "priority": "P0",
        "preconditions": "用户未注册",
        "test_data": "<独立测试账号>",
        "steps": "提交注册",
        "expected_results": "注册成功",
        "notes": "无",
    }
    case.update(overrides)
    return case


class MarkdownValidationTests(unittest.TestCase):
    def test_valid_document(self):
        self.assertEqual(validate(VALID_MARKDOWN), [])

    def test_empty_document_rejected(self):
        errors = validate("# 测试用例\n> 用例数量: 用例 0 条\n")
        self.assertTrue(any("未包含任何有效用例任务项" in error for error in errors))

    def test_missing_required_fields_rejected(self):
        text = VALID_MARKDOWN.replace("      - 前置条件\n        - 用户未注册。\n", "")
        text = text.replace("      - 预期结果\n        - 注册成功。\n", "")
        errors = validate(text)
        self.assertTrue(any("缺少必填字段“前置条件”" in error for error in errors))
        self.assertTrue(any("缺少必填字段“预期结果”" in error for error in errors))

    def test_duplicate_field_rejected(self):
        text = VALID_MARKDOWN.replace(
            "      - 预期结果\n        - 注册成功。\n",
            "      - 预期结果\n        - 注册成功。\n      - 预期结果\n        - 列表可见。\n",
        )
        self.assertTrue(any("重复的“预期结果”字段" in error for error in validate(text)))

    def test_count_uses_case_items(self):
        text = VALID_MARKDOWN.replace("用例 1 条", "用例 2 条")
        self.assertTrue(any("实际用例任务项数量为 1" in error for error in validate(text)))

    def test_orphan_fields_do_not_count_as_cases(self):
        text = "# 测试用例\n> 用例数量: 用例 0 条\n- [ ] 模块\n  - [ ] 功能测试\n    - 预期结果\n      - 孤立字段。\n"
        self.assertTrue(any("未包含任何有效用例任务项" in error for error in validate(text)))

    def test_invalid_type_checked_case_and_duplicate_title_rejected(self):
        text = VALID_MARKDOWN.replace("正向-提交合法注册", "边界-提交合法注册")
        self.assertTrue(any("用例类型非法" in error for error in validate(text)))
        self.assertTrue(any("不得使用已勾选状态" in error for error in validate(VALID_MARKDOWN.replace("- [ ] 正向", "- [x] 正向"))))
        duplicate = VALID_MARKDOWN.replace("用例 1 条", "用例 2 条") + VALID_MARKDOWN.split("- [ ] 注册模块\n", 1)[1].replace("注册模块", "其他模块")
        self.assertTrue(any("用例标题“提交合法注册”重复" in error for error in validate(duplicate)))

    def test_list_format_and_fix(self):
        text = VALID_MARKDOWN.replace("        - 提交注册。\n", "        - 打开页面。\n        - 提交注册。\n")
        self.assertTrue(any("必须全部使用有序列表" in error for error in validate(text)))
        self.assertEqual(validate(fix(text)), [])

    def test_markdown_cli_and_utf8(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "cases.md"
            target.write_text(VALID_MARKDOWN, encoding="utf-8")
            result = subprocess.run(
                [sys.executable, str(ROOT / "scripts" / "validate_markdown_lists.py"), str(target)],
                capture_output=True,
                text=True,
                encoding="utf-8",
                check=False,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            target.write_text(VALID_MARKDOWN, encoding="gbk")
            result = subprocess.run(
                [sys.executable, str(ROOT / "scripts" / "validate_markdown_lists.py"), str(target)],
                capture_output=True,
                text=True,
                encoding="utf-8",
                check=False,
            )
            self.assertNotEqual(result.returncode, 0)


class ExcelTests(unittest.TestCase):
    def test_build_and_validate_default_workbook(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "cases.xlsx"
            build_workbook("测试用例", [excel_case()]).save(target)
            self.assertEqual(validate_excel_file(target), [])
            workbook = load_workbook(target)
            self.assertEqual(workbook.sheetnames, ["测试用例", "生成说明"])
            self.assertEqual(workbook["生成说明"]["B4"].value, "1")

    def test_custom_sheet_validates_with_same_name(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "custom.xlsx"
            build_workbook("测试用例", [excel_case()], sheet_name="自定义").save(target)
            self.assertTrue(validate_excel_file(target))
            self.assertEqual(validate_excel_file(target, sheet_name="自定义"), [])

    def test_formula_like_input_remains_plain_text(self):
        markers = ["=1+1", "+1+1", "-1+1", "@SUM(1)"]
        case = excel_case(requirement_id=markers[0], id=markers[1], title=markers[2], type=markers[3])
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "text.xlsx"
            build_workbook("=标题", [case], meta={"requirement_name": "=需求", "open_questions": "@说明"}).save(target)
            workbook = load_workbook(target, data_only=False)
            sheet = workbook["测试用例"]
            self.assertEqual([sheet.cell(3, column).value for column in range(1, 5)], markers)
            self.assertTrue(all(sheet.cell(3, column).data_type == "s" for column in range(1, 12)))
            self.assertEqual(sheet["A1"].value, "=标题")
            self.assertEqual(sheet["A1"].data_type, "s")
            self.assertEqual(workbook["生成说明"]["B2"].value, "=需求")
            self.assertEqual(workbook["生成说明"]["B6"].value, "@说明")

    def test_excel_required_fields_duplicates_and_priority(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "invalid.xlsx"
            cases = [
                excel_case(priority="P5", preconditions=""),
                excel_case(title="提交合法注册"),
            ]
            build_workbook("测试用例", cases).save(target)
            errors = validate_excel_file(target)
            self.assertTrue(any("缺少必填字段: '前置条件'" in error for error in errors))
            self.assertTrue(any("用例 ID 重复" in error for error in errors))
            self.assertTrue(any("用例标题重复" in error for error in errors))
            self.assertTrue(any("优先级无效" in error for error in errors))

    def test_excel_cli_round_trip(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "cases.json"
            target = Path(directory) / "cases.xlsx"
            source.write_text(
                json.dumps({"meta": {"sheet_name": "自定义"}, "cases": [excel_case()]}, ensure_ascii=False),
                encoding="utf-8",
            )
            subprocess.run(
                [sys.executable, str(ROOT / "scripts" / "create_excel_cases.py"), "--output", str(target), "--input-json", str(source)],
                check=True,
                capture_output=True,
            )
            self.assertEqual(validate_excel_file(target, sheet_name="自定义"), [])

    def test_bundled_excel_template(self):
        self.assertEqual(validate_excel_file(ROOT / "assets" / "excel-test-case-template.xlsx"), [])


if __name__ == "__main__":
    unittest.main()
