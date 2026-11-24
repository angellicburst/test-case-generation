# test-case-generation

一套用于生成 Markdown / Excel 测试用例的可复用工作流与规范，适用于 AI 助手、Agent 或人工执行。

本 README 面向使用者和维护者，只做导航与上手说明；具体执行流程以 [SKILL.md](SKILL.md) 为准，设计与格式规则以 `references/` 下的文档为准。

## 适用场景

- 已提供需求材料，需要拆分功能点并生成可执行、可验证、可追溯的测试用例。
- 需要输出 Markdown 或 Excel（`.xlsx`）格式的用例文件。
- 同一次需求需要同时输出多种格式（共享业务测试设计基线，各格式独立编排与渲染）。

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

用户未指定格式时默认 Markdown。输出路径优先级：用户指定目录或完整文件名 → 项目已有 `03-TCD/{上线日期}/{格式}/` 约定 → 询问用户。`{格式}` 取 `Markdown` 或 `Excel`。`{上线日期}` 为 `YYYY-MM-DD`；用户确认不使用具体日期时，默认文件名、标题和矩阵中写 `未指定日期`；用户明确指定完整文件名时原样使用，且不据此推断上线日期。

## 目录结构

```text
test-case-generation/
├── SKILL.md                    # 执行规范与流程
├── README.md                   # 本文件
├── requirements.txt            # Python 依赖
├── references/
│   ├── test-rules.md           # 用例设计、优先级、测试类型等通用规则
│   ├── traceability-matrix.md  # 需求追溯矩阵结构与校验方式
│   ├── markdown-format.md      # Markdown 字段、层级与校验规则
│   └── excel-format.md         # Excel 字段、工作簿布局与生成脚本说明
├── assets/
│   ├── md-test-case-template.md
│   ├── excel-test-case-template.xlsx
│   └── example-cases-baseline.json
├── scripts/
│   ├── create_excel_cases.py   # 按合并 IR 生成 Excel 工作簿
│   ├── render_cases.py         # 从合并 IR 一键渲染 Markdown / Excel
│   ├── traceability_matrix.py  # 需求追溯条目校验与关联同步
│   ├── validation_utils.py     # 生成器与校验器共享的校验工具
│   ├── validate_markdown_lists.py
│   ├── validate_excel_cases.py
│   ├── validate_traceability_bundle.py
│   └── sync_artifacts.py       # 人工改 Markdown 后回写两种产物
├── tests/
│   └── test_scripts.py         # 脚本自动回归测试
└── evals/
    └── evals.json              # 回归评测用例
```

## 使用方式

把需求材料、上线日期和输出格式交给执行者（AI 助手、Agent 或人工），例如：

> 根据 `01-PRD/xxx.md` 生成 Excel 测试用例，上线日期 2026-11-19。

执行者按 [SKILL.md](SKILL.md) 的流程执行：确认材料完整性 → 建立需求追溯矩阵 → 设计用例 → 矩阵定稿 → 校验产物 → 汇报用例、矩阵和待确认项。

### 直接使用脚本

以下命令中的 `{本目录}` 指本 README 所在目录，不要按当前项目工作目录拼接。

Markdown 校验：

```bash
python "{本目录}/scripts/validate_markdown_lists.py" <文件路径> --traceability-matrix "<矩阵.json>"
python "{本目录}/scripts/validate_markdown_lists.py" <文件路径> --traceability-matrix "<矩阵.json>" --fix
```

`--fix` 只会把多项字段的直接子列表改写为从 1 开始的有序列表，不会修正用例类型、数量声明、层级问题或嵌套说明；写入使用同目录临时文件替换。

Excel 生成与校验：

```bash
python "{本目录}/scripts/create_excel_cases.py" --output "<输出文件.xlsx>" --input-json "<cases.json>" --traceability-matrix "<矩阵.json>"

python "{本目录}/scripts/validate_excel_cases.py" "<文件.xlsx>" --traceability-matrix "<矩阵.json>"
python "{本目录}/scripts/validate_excel_cases.py" "<文件.xlsx>" --traceability-matrix "<矩阵.json>" --sheet-name "<自定义工作表名称>"
```

`create_excel_cases.py` 默认拒绝覆盖已存在的输出文件；用户确认覆盖后，在生成命令末尾追加 `--force`。其 JSON 结构（`meta` 字段、数组字段写法、`sheet_name` 覆盖等）见 [Excel 规范](references/excel-format.md#生成脚本推荐)。

同时交付 Markdown 和 Excel 时，可先执行跨格式总校验；人工改完 Markdown 后，用同步脚本补号并回写两种产物：

```bash
python "{本目录}/scripts/validate_traceability_bundle.py" --markdown "<文件.md>" --excel "<文件.xlsx>"

python "{本目录}/scripts/sync_artifacts.py" --from-markdown "<文件.md>" --excel "<文件.xlsx>" --assign-ids
```

`--coverage-only` 跳过格式细节，仅核对两种格式的用例与追溯是否一致，适合已完成单格式校验后单独执行。
以上命令均为单行，Bash 和 PowerShell 均可直接执行。

## 运行环境

- Python 3.10+（最低要求 3.10），按本机命令使用 `python` 或 `python3`。
- 生成或校验 Excel 需要 `openpyxl`：

  ```bash
  pip install -r "{本目录}/requirements.txt"
  ```

- Markdown 用例文件必须使用 UTF-8 编码。
- 脚本的中文错误信息统一以 UTF-8 输出；接收错误输出的终端或管道应按 UTF-8 解码。

## 关键规则速览

- 通用设计规则（功能点拆分、测试类型、优先级、边界、断言、数据隔离等）见 [测试规则](references/test-rules.md)。
- Markdown 当前格式要求：`- [ ]` 任务项、用例标题下首个非空行使用 `> [需求追溯] 来源：…｜定位：…｜规则：…` 引用、`> 用例数量: 用例 N 条` 数量声明、用例类型仅 `正向` / `反向`、不得使用 `- [x]` 或 `####` 旧式标题。详见 [Markdown 规范](references/markdown-format.md)。
- Excel 要求：`测试用例` 与 `生成说明` 两个工作表、11 个标准列头、优先级 `P0`–`P4`、用例 ID 与标题唯一；生成的 `meta` 必须提供 `requirement_name`、`release_date`，不得填写 `case_count`，`生成说明` 中的用例数量按实际行数自动写入。详见 [Excel 规范](references/excel-format.md)。
- 追溯矩阵是用例的长期伴随交付物；其中用 `待覆盖`、`已覆盖`、`待确认` 管理覆盖状态。`需求/追踪 ID` 仅填写材料原文的正式编号，未提供则留空。

## 常见问题

**提示缺少 openpyxl**

执行 `pip install -r "{本目录}/requirements.txt"`；如果当前环境不允许安装，可先只输出 Markdown，或按 SKILL.md 的预检流程与用户确认后跳过 Excel。

**找不到校验脚本**

确认使用的是 `{本目录}/scripts/...`，而不是当前项目目录下的 `scripts/...`。本目录是 `SKILL.md` 所在目录。

**Markdown 校验失败**

优先检查：是否有 `> 用例数量: 用例 N 条`、用例项是否使用 `- [ ]`、用例类型是否为 `正向` / `反向`、层级是否至少包含功能模块和测试类型两级。多项字段列表编号问题可用 `--fix` 处理。

**终端输出中文乱码**

校验、追溯总校验、同步和 Excel 生成脚本的错误输出均为 UTF-8。若通过管道或日志系统接收错误输出，请将接收端编码设为 UTF-8；这不影响文件内容和校验结果。

## 维护约定

- `SKILL.md` 负责流程，`references/` 负责规则与格式，`scripts/` 负责执行；修改规则时同步检查这三处，避免同一规则多处重复。
- `SKILL.md` 的“交付前检查”是对关键规则的集中复核，规则口径仍以 `references/` 为准。
- 修改脚本或格式后，至少验证一份当前格式 Markdown 用例、Excel 模板、脚本生成的 Excel 样例；代理行为评测按 `evals/evals.json` 逐条执行并人工复核。
- README 只维护导航、上手和常见问题。

### 自动回归测试

```bash
python -m unittest discover -s "{本目录}/tests" -v
```

测试覆盖 Markdown 与 Excel 的正反向校验、自定义工作表、纯文本写入及命令行调用。`evals/evals.json` 中的需求生成与平台策略场景仍需执行对应代理任务并人工复核业务语义，不能用脚本测试代替。
