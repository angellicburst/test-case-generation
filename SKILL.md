---
name: test-case-generation
description: >
  当用户基于需求文档、设计文档、接口说明、原型图或明确的需求描述生成 Markdown 或 Excel (.xlsx) 格式的测试用例时使用本技能；即使用户只说“写用例”但已提供足以拆分测试点的需求材料，也应使用。
  技能生成可执行、可验证、可追溯的本地用例文件。
  不用于仅做需求评审、测试分析、测试计划、缺陷分析，或未提供需求材料的泛泛测试用例咨询。
---

# 测试用例生成

## 运行环境

- Python 3.10+；Markdown 仅需标准库，Excel 另需 `openpyxl`（`pip install -r "{技能根目录}/requirements.txt"`）。
- 按本机命令使用 `python` 或 `python3`。脚本以 UTF-8 读写文件并输出中文，`stdout` / `stderr` 统一 UTF-8。
- Markdown 用例文件必须为 UTF-8。

## 执行流程

### 1. 确认输入与元数据

- 材料整文件不可读，或仅部分可读时，明确指出文件、受影响章节与原因，询问补充或确认忽略；不得静默跳过。
- 识别需求名称、上线日期、测试范围和输出格式；同时提供需求与设计文档时，按[测试规则](references/test-rules.md#多源材料与真实性)的“多源材料基准”处理。
- 未指定格式时默认 Markdown。上线日期标准格式 `YYYY-MM-DD`；无法确定且用户同意时写 `未指定日期`，不得从文件名、目录或当前日期推断。项目无日期概念时，由用户指定文件名或替代标识。
- 缺失上线日期、输出目录等元数据与覆盖确认、待确认问题合并为一次提问；用户希望直接出稿时可先用 `未指定日期` 起草并在汇报中说明。交互细则见[测试规则](references/test-rules.md#元数据交互)。
- 仅当核心业务规则或测试范围缺失、冲突到无法判断时，才暂停设计请求澄清；局部细节缺失按[测试规则](references/test-rules.md)记入 `待确认`。
- 新建生成时不读取历史用例；维护模式才读取用户指定的目标文件与需求变更。

### 2. 分配需求任务

一个独立需求对应一份用例文档，不合并多个需求。输出路径优先级：用户指定目录或完整文件名 → 项目已有 `03-TCD/{上线日期}/{格式}/` 约定（`{格式}` 取 `Markdown` 或 `Excel`）→ 询问用户。默认文件名 `测试用例-{上线日期}-{需求名称}.{md|xlsx}`；用户指定完整文件名时原样使用。

### 3. 需求分析与测试设计

读取[测试规则](references/test-rules.md)，按其“功能点拆分与用例组织”“测试设计能力”“用例设计原则”“执行准则”拆分功能点与测试要点：

- 每条用例只有单一可验证逻辑，具备前置操作、测试步骤和可判定预期。
- 纯发布配置、初始化字典、仅查询加速的普通索引 DDL 不单列运维用例，归并到读写该表的业务功能点。
- 待确认项按风险排序，单模块常规需求收敛在 3~5 项内。
- 同时输出多种格式时，把共享业务设计、追溯关系与各格式所需字段统一写入单一临时 IR，由 `render_cases.py` 确定性渲染各格式；不手写多份同构副本，也不手工对单格式成稿做有损互转。

### 4. 编写临时合并 IR

- 按[需求追溯与待确认规范](references/traceability-matrix.md)建立条目，写入**单一临时 JSON**（结构见 `assets/example-cases-baseline.json`）：`meta`、`traceability`、`cases`。
- 用例 ID 为 `TC-{需求简写}-{3 位流水号}`；需求简写由 LLM 从需求名称核心词提炼 2~8 位大写字母/数字，同一需求文档统一前缀，不绑定模块名。ID 是永久主键：删除允许跳号，新增取现有最大值 +1；仅定稿前允许 `--renumber` 重排。完整规则见[测试规则](references/test-rules.md#用例-id-稳定策略)。
- 规模策略：单模块、预计 ≤30 条或维护更新，默认直接串行编写；多模块大规模 >30 条且平台支持并行时，按功能模块并行设计片段，由主代理合并去重后统一渲染。
- `cases` 每条包含 `id`、`module`、`type`、`group`（可空）、`polarity`（正向/反向）、`title`、`design_method`、`priority`、`pre_operations`、`test_data`、`steps`、`expected_results`、`post_operations`、`traceability_refs`；`design_method` 仅在生成 Excel 时必填，纯 Markdown 可空；`meta.case_count` 不得手填（按 `cases` 自动计算）；已移除 `requirement_id` 与用例级 `notes`。数组每项填纯文本，不写 `- ` 或 `1. ` 前缀；渲染时单项原样输出，多项自动编号。
- 字段取值、枚举与渲染按目标格式的规范和模板确定，见[格式索引](#格式索引)。

### 5. 一键渲染与校验

渲染前做最小预检（只检查本次要用的能力）：

- Excel：`python -c "import openpyxl"`；失败时说明无法生成与校验 Excel，并给出“安装依赖 / 仅输出 Markdown”选项。纯 Markdown 不受影响。
- 校验脚本位于 `{技能根目录}/scripts/`，不得按当前项目工作目录拼接。平台无法执行命令时跳过预检：Markdown 可按规范生成并在汇报中标注“未校验”；Excel 不具备生成条件，应向用户说明。

```bash
python "{技能根目录}/scripts/render_cases.py" --baseline "<临时ir.json>" [--markdown "<输出.md>"] [--excel "<输出.xlsx>"] --cleanup-baseline
```

`--markdown` 与 `--excel` 至少指定一个。脚本统一渲染、全量校验并原子写入，全部通过后删除临时 IR；输出已存在时默认拒绝覆盖，用户确认后追加 `--force`。脚本可执行时，失败须按报错修正 IR 后重跑，不得交付未校验产物。

写入后按[交付前检查](#交付前检查)复核业务语义；格式校验通过不代表业务正确。

### 6. 汇报与维护

- 汇报每种格式的文件名、用例数量、可点击路径，并逐项展示 `待确认` 条目的来源/定位、未明确项、测试假设和确认问题；无待确认项时明确说明。
- 交付目录只保留最终 `.md` / `.xlsx`，不保留临时 IR 或 JSON 矩阵。
- 确认变更：用 `case_baseline.py export` 从现有产物导出 IR，编辑后以 `render_cases.py --assign-ids --force --cleanup-baseline` 重新渲染；或直接人工修改 Markdown 后用 `sync_artifacts.py` 补号、重排并同步两种格式（规则见[测试规则](references/test-rules.md#确认维护)）。
- 确认的规则转 `已覆盖`；无需覆盖的条目及相应用例删除；用例数量、追溯表与引用在两种格式中保持一致。

## 常用命令

```bash
python "{技能根目录}/scripts/render_cases.py" --baseline "<ir.json>" [--markdown "<输出.md>"] [--excel "<输出.xlsx>"] [--assign-ids] [--force] [--cleanup-baseline]
python "{技能根目录}/scripts/case_baseline.py" export [--markdown "<文件.md>"] [--excel "<文件.xlsx>"] --output "<ir.json>" [--force]
python "{技能根目录}/scripts/sync_artifacts.py" --from-markdown "<文件.md>" [--excel "<文件.xlsx>"] [--assign-ids] [--renumber] [--default-design-method "等价类划分"]
python "{技能根目录}/scripts/validate_markdown_lists.py" "<文件.md>" [--baseline "<ir.json>"] [--fix]
python "{技能根目录}/scripts/validate_excel_cases.py" "<文件.xlsx>" [--sheet-name "测试用例"]
python "{技能根目录}/scripts/validate_traceability_bundle.py" [--markdown "<文件.md>"] [--excel "<文件.xlsx>"] [--baseline "<ir.json>"] [--coverage-only]
```

`render_cases.py` 与 `case_baseline.py export` 的 Markdown/Excel 至少指定一个；跨格式校验脚本至少提供一个文件，`--coverage-only` 必须同时提供 `--markdown` 与 `--excel`。

## 格式索引

| 格式 | 规范 | 模板 | 校验脚本 |
| :--- | :--- | :--- | :--- |
| Markdown | [Markdown 规范](references/markdown-format.md) | [Markdown 模板](assets/md-test-case-template.md) | `validate_markdown_lists.py` |
| Excel | [Excel 规范](references/excel-format.md) | [Excel 模板](assets/excel-test-case-template.xlsx) | `validate_excel_cases.py` |

## 交付前检查

- [ ] 每条用例单一逻辑；前置操作、测试步骤完整；预期结果可判定，且可回溯到关联的 `已覆盖` 条目，无臆造的产品行为。
- [ ] 用户给出的并列子条件均已落到用例；边界上点已覆盖，离点紧邻上点（如上限“6 个月”对应“6 个月零 1 天”）。
- [ ] 数据准备使用独立隔离语义的脱敏占位符，不含真实身份、手机号、账号或凭证。
- [ ] 需要恢复环境或清理数据时写入后置操作，未重复塞进测试步骤；撤销/回滚本身是业务功能时除外。
- [ ] 正常、异常、边界与关联回归均有材料依据；无重复或等价用例；回归场景位于 `回归测试` 下。
- [ ] 待确认项已按风险排序并展示，未把未确认假设写成确定预期。
- [ ] 已通过 `render_cases.py` 全量校验（同时渲染双格式时已内置跨格式校验；分步生成或单格式分别校验后需再执行 `validate_traceability_bundle.py --coverage-only` 核对）；经用户确认标注“未校验”的除外。
