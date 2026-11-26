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

- 材料整文件不可读或部分可读时，指出文件、受影响章节与原因，询问补充或确认忽略；不得静默跳过。
- 识别需求名称、上线日期、测试范围和输出格式；调用方已给出的元数据直接采用、不重复询问，委派执行时回传最终值；需求与设计文档同时提供时按[测试规则](references/test-rules.md#多源材料与真实性)的“多源材料基准”处理。
- **元数据决策顺序**：先确认缺失的上线日期与输出目录，再按目标路径查同名文件；其他日期文件不作日期候选、不在定日期前问覆盖；无 `TCD` 约定且用户未指定目录时，合并提问必须确认目录（当前目录可作默认）。
- 未指定格式默认 Markdown。上线日期为 `YYYY-MM-DD` 或 `未指定日期`，不得从文件名、当前日期、目录名或分支名推断（按日期分批的输入目录可作候选并确认）；`20261128`、`2026/11/28` 等直接规范化为 `YYYY-MM-DD` 并在汇报说明，不追问；无日期概念时由用户指定文件名或替代标识。
- 新建不读已有用例；同名文件存在须先确认覆盖（自动模式或 Git 跟踪也不得绕过），维护模式才读目标文件与变更；存在性用 `Path.exists()` 或 `test -e`，失败/报错/乱码按无法确认处理并询问。
- 仅核心规则/范围缺失冲突时暂停澄清；局部细节记入 `待确认`；缺失项合并提问，细则见[测试规则](references/test-rules.md#元数据交互)。

### 2. 分配需求任务

一个独立需求对应一份文档。路径优先级：用户指定目录/文件名 → 项目 `TCD/{上线日期}/{格式}/` 约定 → 询问用户；`{格式}` 取 `Markdown`/`Excel`。默认文件名 `测试用例-{上线日期}-{需求名称}.{md|xlsx}`，用户指定完整文件名时原样使用。

### 3. 需求分析与测试设计

按[测试规则](references/test-rules.md)的“功能点拆分与用例组织”“测试设计能力”“用例设计原则”“执行准则”拆分功能点与测试要点：

- 用例单一逻辑，含前置、步骤与可判定预期。
- 纯发布配置、初始化字典、仅查询加速的普通索引 DDL 不单列运维用例，归并到读写该表的业务功能点。
- 待确认项按风险排序，单模块常规需求收敛在 3~5 项内。
- 多格式输出时，把业务设计、追溯与字段统一写入单一临时 IR，由 `render_cases.py` 确定性渲染；不手写同构副本，不手工做有损互转。

### 4. 编写临时合并 IR

- 按[追溯规范](references/traceability-matrix.md)建条目；**单一临时 JSON**与脚本放系统临时目录（`%TEMP%`/`/tmp`，勿放交付目录或工作区），结构见 `assets/example-cases-baseline.json`。
- 用例 ID 为 `TC-{需求简写}-{3 位流水号}`，是永久主键（删除可跳号、新增取最大值 +1，`--renumber` 仅定稿前）；需求简写取需求名称核心词 2~8 位大写字母/数字，单需求统一前缀、不绑定模块名，完整规则见[测试规则](references/test-rules.md#用例-id-稳定策略)。
- 规模策略：单模块、≤30 条或维护更新串行；多模块 >30 条且平台支持并行时按模块并行设计，由主代理合并去重后统一渲染。
- `cases` 每条包含 `id`、`module`、`type`、`group`（可空）、`polarity`（正向/反向）、`title`、`design_method`、`priority`、`pre_operations`、`test_data`、`steps`、`expected_results`、`post_operations`、`traceability_refs`；`design_method` 仅在生成 Excel 时必填，纯 Markdown 可空；`meta.case_count` 不得手填（按 `cases` 自动计算）。数组每项填纯文本，不写 `- ` 或 `1. ` 前缀；渲染时单项原样输出，多项自动编号。
- 起草态 IR：`schema_version` 固定 `"2.0"`；`meta` 含 `requirement_name`、`release_date`、`coverage`/`notes`（可空）、`sheet_name`（默认 `"测试用例"`），`title` 可省由脚本拼接；`traceability` 的 `已覆盖`/`待覆盖` 条目不手填 `case_ids`（由 `traceability_refs` 派生）。
- 字段取值、枚举与渲染按目标格式的规范和模板确定，见[格式索引](#格式索引)。

### 5. 一键渲染与校验

渲染前先按[交付前检查](#交付前检查)除末条全量校验外的全部业务自检项自检 IR 业务语义，再做最小能力预检（只查本次所需）：

- 能力异常分级（脚本在 `{技能根目录}/scripts/`）：① 缺 `openpyxl` → 不可生成/校验 Excel，给出“安装依赖 / 仅出 Markdown”；② 校验脚本缺失 → 报路径与影响，确认后标注“未校验”继续（不得伪称通过）；③ 无法执行命令 → 跳过预检，MD 标注“未校验”，Excel 说明不具备生成条件。
- 执行[常用命令](#常用命令)的 `render_cases.py`：统一渲染、校验、原子写入；自检通过且无待确认、无后续确认时加 `--cleanup-baseline`，否则保留 IR 至确认完成；输出已存在须确认后加 `--force`（同轮重渲染除外）。失败须修正 IR 重跑；Excel/WPS 占用时脚本已回滚并保留 IR，关闭后原样重跑，严禁拆写单格式或留候选文件。

`render_cases.py` 已内置 Markdown、Excel 与跨格式全量校验，成功后不再单跑 `validate_markdown_lists.py`/`validate_excel_cases.py`；渲染前预检起草态用 `case_baseline.py check --draft`（仅证明可渲染，不替代默认终态 `check`）。写入后用 `case_baseline.py inspect` 抽查（不裸写 `python -c`）；校验通过不代表业务正确，发现问题按第 6 步改 IR 重渲染再交付。

### 6. 汇报与维护

- 汇报每种格式的文件名、数量与可点击路径，逐项展示 `待确认` 的来源/定位、未明确项、测试假设与确认问题；无待确认项时明确说明。
- 交付目录只保留最终 `.md` / `.xlsx`，不保留临时 IR 或 JSON 矩阵。
- 确认变更：IR 仍在则直接改；**改后必须重走第 5 步渲染前自检与 `render_cases.py --assign-ids --force` 重渲染，再对全部修改/新增用例及涉及追溯条目执行 `case_baseline.py inspect --ids ... --trace-ids ...` 抽查闭环**；有待确认/待复核分歧而留 IR 时，不得清零交付物中的 `待确认`。确认完成后再加 `--cleanup-baseline`。无 IR（已清理或跨会话/换模型无法定位）时，用 `case_baseline.py export` 导出到临时目录再改再渲染；微调 Markdown/增删用 `sync_artifacts.py`（见[测试规则](references/test-rules.md#确认维护)）。
- 确认规则转 `已覆盖`；无需覆盖的条目及用例删除；两种格式的用例数量、追溯表与引用保持一致。

## 常用命令

```bash
python "{技能根目录}/scripts/render_cases.py" --baseline "<ir.json>" [--markdown "<输出.md>"] [--excel "<输出.xlsx>"] [--assign-ids] [--force] [--cleanup-baseline]
python "{技能根目录}/scripts/case_baseline.py" export [--markdown "<文件.md>"] [--excel "<文件.xlsx>"] --output "<ir.json>" [--force]
python "{技能根目录}/scripts/case_baseline.py" check --baseline "<ir.json>" [--draft] [--assign-ids] [--require-excel-fields]
python "{技能根目录}/scripts/case_baseline.py" inspect [--markdown "<文件.md>"] [--excel "<文件.xlsx>"] [--baseline "<ir.json>"] [--ids ID1,ID2] [--trace-ids T1,T2]
python "{技能根目录}/scripts/sync_artifacts.py" --from-markdown "<文件.md>" [--excel "<文件.xlsx>"] [--assign-ids] [--renumber] [--default-design-method "等价类划分"]
python "{技能根目录}/scripts/validate_markdown_lists.py" "<文件.md>" [--baseline "<ir.json>"] [--fix]
python "{技能根目录}/scripts/validate_excel_cases.py" "<文件.xlsx>" [--sheet-name "测试用例"]
python "{技能根目录}/scripts/validate_traceability_bundle.py" [--markdown "<文件.md>"] [--excel "<文件.xlsx>"] [--baseline "<ir.json>"] [--coverage-only]
```

`render_cases.py`、`export`、`inspect` 与跨格式校验至少指定一源；抽查产物优先 `inspect`；`check` 默认校验终态，`--draft` 只读预检起草态（`--assign-ids` 仅可与 `--draft` 连用）；`--coverage-only` 须同传双格式。

## 格式索引

| 格式 | 规范 | 模板 | 校验脚本 |
| :--- | :--- | :--- | :--- |
| Markdown | [Markdown 规范](references/markdown-format.md) | [Markdown 模板](assets/md-test-case-template.md) | `validate_markdown_lists.py` |
| Excel | [Excel 规范](references/excel-format.md) | [Excel 模板](assets/excel-test-case-template.xlsx) | `validate_excel_cases.py` |

## 交付前检查

- [ ] 每条用例单一逻辑；前置、步骤完整；预期可判定，`rule` 的条件与结果在预期中有对应断言，且可回溯到 `已覆盖` 条目；无臆造产品行为，材料/确认外枚举、配置、表名、字段值用 `<...>` 占位。
- [ ] 用户并列子条件均已落到用例；边界上点已覆盖，离点紧邻上点（如上限“6 个月”对应“6 个月零 1 天”）。
- [ ] 数据准备使用独立隔离语义的脱敏占位符，不含真实身份、手机号、账号或凭证。
- [ ] 需恢复环境或清理数据时写入后置操作，不重复塞进测试步骤；撤销/回滚本身是业务功能时除外。
- [ ] 正常、异常、边界与关联回归均有材料依据；无重复或等价用例；各判定分支独立成例，不得藏入回归；`回归测试` 仅放受影响的既有行为。
- [ ] 待确认项按风险排序展示；未确认假设不写成确定预期；确认口径优先，前置/步骤/预期不含确认排除的依赖。
- [ ] 原 PRD 已覆盖条目 `source`/`location`/`rule` 未被改写（仅情形 3 失去全部支撑用例转 `待确认` 时允许将原 `rule` 摘要迁入 `issue` 并留痕）、确认日期为当日实际日期、待复核分歧不清空 `待确认`（见[追溯规范](references/traceability-matrix.md#待确认项的记录与收敛)）。
- [ ] 确认修改后标题、前置、步骤观测、预期、后置与追溯引用闭环（见[测试规则](references/test-rules.md#确认维护)）。
- [ ] 模块按独立进程/接口/页面/定时任务拆分，不按内部逻辑拆分、不以测试类型作模块（见[测试规则](references/test-rules.md#功能点拆分与用例组织)）。
- [ ] 已通过 `render_cases.py` 全量校验（双格式渲染内置跨格式校验；双格式分步渲染须补跑 `validate_traceability_bundle.py --coverage-only`）；经确认标注“未校验”的除外。
