# test-case-generation

一套用于生成 Markdown / Excel 测试用例的可复用工作流与规范，适用于 AI 助手、Agent 或人工执行。

本 README 面向使用者和维护者，只做导航与上手说明；具体执行流程以 [SKILL.md](SKILL.md) 为准，设计与格式规则以 `references/` 下的文档为准。

## 适用场景

- 已提供需求材料，需要拆分功能点并生成可执行、可验证、可追溯的测试用例。
- 需要输出 Markdown 或 Excel（`.xlsx`）格式的用例文件。
- 同一次需求需要同时输出多种格式（每种格式分别从原始材料重新设计，不互相转换）。

不适用于：

- 仅做需求评审、测试分析、测试计划或缺陷分析。
- 没有需求材料的泛泛测试用例咨询。

## 输入与输出

**输入**：需求文档、设计文档、接口说明、原型图或明确的需求描述；可选指定上线日期、输出格式和输出目录。

**输出**：

| 格式 | 默认文件名 | 说明 |
| :--- | :--- | :--- |
| Markdown | `测试用例-{上线日期}-{需求名称}.md` | 任务清单树，便于评审和版本管理 |
| Excel | `测试用例-{上线日期}-{需求名称}.xlsx` | 11 列标准表格，含 `测试用例` 与 `生成说明` 工作表 |

用户未指定格式时默认 Markdown。输出路径优先级：用户指定目录或文件名 → 项目已有 `03-TCD/{上线日期}/{格式}/` 约定 → 询问用户。`{格式}` 取 `Markdown` 或 `Excel`。

## 目录结构

```text
test-case-generation/
├── SKILL.md                    # 执行规范与流程
├── README.md                   # 本文件
├── references/
│   ├── test-rules.md           # 用例设计、优先级、测试类型等通用规则
│   ├── markdown-format.md      # Markdown 字段、层级与校验规则
│   └── excel-format.md         # Excel 字段、工作簿布局与生成脚本说明
├── assets/
│   ├── md-test-case-template.md
│   └── excel-test-case-template.xlsx
├── scripts/
│   ├── create_excel_cases.py   # 按 JSON 生成 Excel 工作簿
│   ├── validate_markdown_lists.py
│   └── validate_excel_cases.py
└── evals/
    └── evals.json              # 回归评测用例
```

## 使用方式

把需求材料、上线日期和输出格式交给执行者（AI 助手、Agent 或人工），例如：

> 根据 `01-PRD/xxx.md` 生成 Excel 测试用例，上线日期 2026-11-19。

执行者按 [SKILL.md](SKILL.md) 的流程执行：确认材料完整性 → 拆分需求与功能点 → 设计用例 → 生成并校验产物 → 汇报用例数量、文件路径和待确认项。

### 直接使用脚本

以下命令中的 `{本目录}` 指本 README 所在目录，不要按当前项目工作目录拼接。

Markdown 校验：

```bash
python "{本目录}/scripts/validate_markdown_lists.py" <文件路径>
python "{本目录}/scripts/validate_markdown_lists.py" <文件路径> --fix
```

`--fix` 只会把多项字段列表改写为从 1 开始的有序列表，不会修正用例类型、数量声明或层级问题。

Excel 生成与校验：

```bash
python "{本目录}/scripts/create_excel_cases.py" --output "<输出文件.xlsx>" --input-json "<cases.json>"

python "{本目录}/scripts/validate_excel_cases.py" "<文件.xlsx>"
```

`create_excel_cases.py` 的 JSON 结构（`meta` 字段、数组字段写法、`sheet_name` 覆盖等）见 [Excel 规范](references/excel-format.md#生成脚本推荐)。

## 运行环境

- Python 3.8 及以上（已在 3.12 实测通过），按本机命令使用 `python` 或 `python3`。
- 生成或校验 Excel 需要 `openpyxl`：

  ```bash
  pip install openpyxl
  ```

- Markdown 用例文件必须使用 UTF-8 编码。
- Windows 终端中文乱码时，可先执行 `chcp 65001`，或设置 `PYTHONIOENCODING=utf-8`。

## 关键规则速览

- 通用设计规则（功能点拆分、测试类型、优先级、边界、断言、数据隔离、待确认清单等）见 [测试规则](references/test-rules.md)。
- Markdown 当前格式要求：`- [ ]` 任务项、`> 用例数量: 用例 N 条` 数量声明、用例类型仅 `正向` / `反向`、不得使用 `- [x]` 或 `####` 旧式标题。详见 [Markdown 规范](references/markdown-format.md)。
- Excel 要求：`测试用例` 与 `生成说明` 两个工作表、11 个标准列头、优先级 `P0`–`P4`、用例 ID 与标题唯一。详见 [Excel 规范](references/excel-format.md)。
- 需求存在未定义细节时，执行者会生成《需求存疑与待确认清单》；确认后反向优化用例并删除清单（Markdown 删章节，Excel 清空 `生成说明` 对应字段）。
- 只接受 `> 用例数量: 用例 N 条` 的数量声明；`description: 用例 N 条` 等历史写法不再被接受，历史文件需要先改为当前格式，或按当前需求重新生成。

## 常见问题

**提示缺少 openpyxl**

执行 `pip install openpyxl`；如果当前环境不允许安装，可先只输出 Markdown，或按 SKILL.md 的预检流程与用户确认后跳过 Excel。

**找不到校验脚本**

确认使用的是 `{本目录}/scripts/...`，而不是当前项目目录下的 `scripts/...`。本目录是 `SKILL.md` 所在目录。

**Markdown 校验失败**

优先检查：是否有 `> 用例数量: 用例 N 条`、用例项是否使用 `- [ ]`、用例类型是否为 `正向` / `反向`、层级是否至少包含功能模块和测试类型两级。多项字段列表编号问题可用 `--fix` 处理。

**终端输出中文乱码**

执行 `chcp 65001`，或设置 `PYTHONIOENCODING=utf-8`；这不影响文件内容和校验结果。

## 维护约定

- `SKILL.md` 负责流程，`references/` 负责规则与格式，`scripts/` 负责执行；修改规则时同步检查这三处，避免同一规则多处重复。
- 修改脚本或格式后，至少验证一份当前格式 Markdown 用例、Excel 模板、脚本生成的 Excel 样例，并跑一遍 `evals/evals.json` 中的回归场景。
- README 只维护导航、上手和常见问题；具体规则变更以 `SKILL.md` 和 `references/` 为准。
