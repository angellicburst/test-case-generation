# test-case-generation

一套用于生成 Markdown / Excel 测试用例的可复用工作流与规范，适用于 AI 助手、Agent 或人工执行。

本 README 面向使用者和维护者，只做导航与上手说明；执行流程以 [SKILL.md](SKILL.md) 为准，设计规则以 [references/test-rules.md](references/test-rules.md) 为准，格式以 `references/` 下对应文档为准。

## 适用场景

- 已提供需求材料，需要拆分功能点并生成可执行、可验证、可追溯的测试用例。
- 需要输出 Markdown 或 Excel（`.xlsx`）格式的用例文件。
- 同一次需求需要同时输出多种格式（统一写入单一临时 IR，由脚本确定性渲染）。

不适用于仅做需求评审、测试分析、测试计划或缺陷分析，以及没有需求材料的泛泛测试用例咨询。

## 输入与输出

**输入**：需求文档、设计文档、接口说明、原型图或明确的需求描述；可选指定上线日期、输出格式和输出目录。

**输出**：

| 格式 | 默认文件名 | 说明 |
| :--- | :--- | :--- |
| Markdown | `测试用例-{上线日期}-{需求名称}.md` | 任务清单树，含 5 个内容字段与文档内需求追溯 |
| Excel | `测试用例-{上线日期}-{需求名称}.xlsx` | 11 列用例表 + `生成说明` + `需求追溯`，有待确认项时增加 `待确认项` 工作表 |

用户未指定格式时默认 Markdown。输出路径优先级：用户指定目录或完整文件名 → 项目已有 `TCD/{上线日期}/{格式}/` 约定 → 询问用户。`{格式}` 取 `Markdown` 或 `Excel`。`{上线日期}` 为 `YYYY-MM-DD`；用户确认不使用具体日期时写 `未指定日期`。

**交付目录只保留最终的 `.md` / `.xlsx`**：追溯与待确认数据已内嵌在两种格式中，不再生成矩阵 JSON 或临时 IR；临时 IR 在渲染成功后由 `--cleanup-baseline` 删除。

## 目录结构

```text
test-case-generation/
├── SKILL.md                       # 执行流程与交付前检查
├── README.md                      # 本文件
├── requirements.txt               # Python 依赖
├── references/
│   ├── test-rules.md              # 测试设计、优先级、ID、待确认等业务规则
│   ├── traceability-matrix.md     # 需求追溯与待确认规范
│   ├── markdown-format.md         # Markdown 结构、字段与校验
│   └── excel-format.md            # Excel 11 列与 3/4 工作表布局
├── assets/
│   ├── md-test-case-template.md
│   ├── excel-test-case-template.xlsx
│   └── example-cases-baseline.json  # 合并 IR 结构示例
├── scripts/
│   ├── case_baseline.py           # IR 校验、导出、抽查、补号、重排
│   ├── render_cases.py            # 从临时 IR 一键渲染 Markdown / Excel
│   ├── render_markdown_cases.py   # Markdown 渲染器
│   ├── create_excel_cases.py      # Excel 工作簿构建器
│   ├── sync_artifacts.py          # 人工改 Markdown 后同步两种产物
│   ├── traceability_matrix.py     # 追溯条目校验与关联汇总
│   ├── validation_utils.py        # 生成器与校验器共享工具
│   ├── validate_markdown_lists.py
│   ├── validate_excel_cases.py
│   └── validate_traceability_bundle.py
├── tests/
│   └── test_scripts.py            # 脚本自动回归测试
└── evals/
    └── evals.json                 # 代理行为评测用例
```

## 使用方式

把需求材料、上线日期和输出格式交给执行者（AI 助手、Agent 或人工），例如：

> 根据 `需求文档.md` 生成 Excel 测试用例，上线日期 2026-11-19。

执行者按 [SKILL.md](SKILL.md) 执行：确认输入与元数据 → 分配需求 → 分析与设计 → 编写临时合并 IR → 一键渲染与校验 → 汇报与维护。

### 直接使用脚本

以下命令中的 `{本目录}` 指本 README 所在目录，不要按当前项目工作目录拼接；均为单行命令。

从临时 IR 生成两种格式（推荐入口）：

```bash
python "{本目录}/scripts/render_cases.py" --baseline "<ir.json>" --markdown "<输出.md>" --excel "<输出.xlsx>" --cleanup-baseline
python "{本目录}/scripts/render_cases.py" --baseline "<ir.json>" --markdown "<输出.md>" --force
```

从现有产物导出 IR 以便维护：

```bash
python "{本目录}/scripts/case_baseline.py" export --markdown "<文件.md>" --excel "<文件.xlsx>" --output "<ir.json>" --force
python "{本目录}/scripts/case_baseline.py" check --baseline "<ir.json>"
```

抽查现有产物的用例分布与单条详情（只读，不改文件）：

```bash
python "{本目录}/scripts/case_baseline.py" inspect --markdown "<文件.md>"
python "{本目录}/scripts/case_baseline.py" inspect --excel "<文件.xlsx>" --ids TC-XXX-001,TC-XXX-002
```

人工改完 Markdown 后补号、同步 Excel（存在缺失设计方法的用例时需提供，如新增用例或首次由 Markdown 生成 Excel）：

```bash
python "{本目录}/scripts/sync_artifacts.py" --from-markdown "<文件.md>" --excel "<文件.xlsx>" --assign-ids --default-design-method "等价类划分"
python "{本目录}/scripts/sync_artifacts.py" --from-markdown "<文件.md>" --renumber
```

单独校验：

```bash
python "{本目录}/scripts/validate_markdown_lists.py" "<文件.md>" [--baseline "<ir.json>"] [--fix]
python "{本目录}/scripts/validate_excel_cases.py" "<文件.xlsx>" [--sheet-name "测试用例"]
python "{本目录}/scripts/validate_traceability_bundle.py" --markdown "<文件.md>" --excel "<文件.xlsx>" [--baseline "<ir.json>"] [--coverage-only]
```

`render_cases.py` 与 `sync_artifacts.py` 都会先渲染到临时文件、全量校验，再备份并原子替换；任一写入失败会回滚并提示文件占用。`render_cases.py` 检测到输出已存在时默认拒绝覆盖，用户确认后追加 `--force`；`sync_artifacts.py` 属维护模式，用户提出更新即授权覆盖指定文件。

**Excel 被占用时**：脚本会提示“文件被占用”并回滚，临时 IR 保留在原处。关闭 Excel/WPS 后原样重跑同一条命令即可，不要改成只写 Markdown 或另存候选文件——那会让两种格式不同步。只读文件与目录权限不足会分别提示“文件不可写”与“输出目录不可写”。

## 运行环境

- Python 3.10+，按本机命令使用 `python` 或 `python3`。
- Markdown 仅需标准库；生成或校验 Excel 需要 `openpyxl`：

  ```bash
  pip install -r "{本目录}/requirements.txt"
  ```

- Markdown 文件必须为 UTF-8；脚本以 UTF-8 读写并输出中文，`stdout` / `stderr` 统一 UTF-8。

## 关键规则速览

- 通用设计规则（功能点拆分、测试类型、优先级、边界、断言、数据隔离、ID 稳定、待确认收敛等）见 [测试规则](references/test-rules.md)。
- Markdown：`- [ ]` 任务树；用例标题只写 `{正向|反向}-{标题}-{P0..P4}`；下一行是 `> [TC-ID] 追溯: TRACE-ID`；5 个内容字段按 前置操作 → 测试数据 → 测试步骤 → 预期结果 → 后置操作 排列；文末 `## 需求追溯` / `### 需求追溯表`，有待确认项时增加 `### 待确认项`。详见 [Markdown 规范](references/markdown-format.md)。
- Excel：11 列固定表头（追溯 ID、用例 ID、用例标题、用例类型、设计方法、优先级、前置操作、测试数据、测试步骤、预期结果、后置操作），`测试用例` / `生成说明` / `需求追溯` 三个工作表，有待确认项时增加 `待确认项`；仅纯文本，不含公式。详见 [Excel 规范](references/excel-format.md)。
- 用例 ID 为 `TC-{需求简写}-{3 位流水号}`，是永久主键：删除可跳号，新增取最大值 +1；仅定稿前可 `--renumber`。
- 追溯条目用 `待覆盖`、`已覆盖`、`待确认` 管理，内嵌在产物中并与用例双向关联；交付前不得残留 `待覆盖`。详见 [需求追溯与待确认规范](references/traceability-matrix.md)。

## 常见问题

**提示缺少 openpyxl**

执行 `pip install -r "{本目录}/requirements.txt"`；若环境不允许安装，可只输出 Markdown。Excel 必须安装 `openpyxl` 才能生成与校验。

**找不到校验脚本**

确认使用的是 `{本目录}/scripts/...`，而不是当前项目目录下的 `scripts/...`。本目录是 `SKILL.md` 所在目录。

**Markdown 校验失败**

优先检查：一级标题与 `> 用例数量: 用例 N 条` 是否存在且数量一致；用例项是否使用 `- [ ]`；类型是否为 `正向` / `反向`；层级是否为模块 → 测试类型 →（可选分组）→ 用例；5 个字段顺序与必填项是否完整；引用行是否为 `> [TC-ID] 追溯: ...` 且顺序与追溯表一致。多项字段编号可用 `--fix` 处理。

**Excel 校验失败**

优先检查：工作表是否齐全、主表是否恰好 `A1:K1` 合并、表头是否为 11 列固定顺序、是否冻结 `A3`、筛选范围是否正确、是否存在公式、`生成说明` 的 `用例数量` 是否与实际行数一致、追溯表是否至少 1 行数据。

**写入时提示文件被占用**

Excel/WPS 打开着目标文件时，Windows 会阻止替换。关闭程序后原样重跑同一条命令；脚本已回滚且保留了临时 IR，不需要重新生成。用 `case_baseline.py inspect` 可先确认当前产物内容是否已是期望版本。

**终端输出中文乱码**

脚本输出均为 UTF-8；通过管道或日志系统接收时，请将接收端编码设为 UTF-8，这不影响文件内容和校验结果。

## 维护约定

- `SKILL.md` 负责流程，`references/` 负责规则与格式，`scripts/` 负责执行；同一规则只保留一个正文出处，其他位置用链接指向。
- 修改脚本或格式后，至少验证一份 Markdown 用例、Excel 模板与脚本生成的 Excel 样例；代理行为评测按 `evals/evals.json` 逐条执行并人工复核业务语义。

### 自动回归测试

```bash
python -m unittest discover -s "{本目录}/tests" -v
```
