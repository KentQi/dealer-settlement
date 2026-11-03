# Changelog

本项目的所有重要变更记录于此，格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，
版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

## [Unreleased]

## [1.3.0] - 2026-11-03

修复确认单流程的用户体验缺陷：原实现要求用户把下载的 `确认单.json` 手动搬进客户目录，普通用户不具操作前提。现在改为**零文件操作**——确认页提交后，助手自动定位并归档，用户全程不需要接触/移动文件。

### Added

- **确认单自动定位（`locate_confirm.py`）**：助手在客户目录（含 `_确认记录/`）与平台常见下载/桌面目录中查找确认单（文件名前缀 `合同确认单_*.json`，兼容旧版日期命名）；校验 `schema: contract-confirm/v1`、`meta.operator` 非空、`fields` 数组；`--json` 输出候选列表，退出码 0（唯一命中）/ 2（多份，给出最新一份与全部候选）/ 1（未找到并说明查过的目录）；只读，不移动/不修改任何文件。
- **确认页零文件操作（`make_confirm_html.py`）**：
  - 成功后弹窗重写为普通用户文案：「已提交成功。回到对话，对助手说一句『好了』即可。」并明确「助手会自动找到并归档，不需要你移动任何文件」。
  - 支持 File System Access API 的浏览器（Chrome / Edge）新增「存到客户目录…」按钮，可直接把确认单写入用户选择的客户目录；不支持时自动回退到浏览器下载，不影响主流程。
  - 文件名改为确定性前缀 `合同确认单_<客户名>_<提交时间>.json`，便于自动定位；客户名做文件系统安全化。
  - 移除所有「把 JSON 放进目录 / 复制 JSON 到目录」类引导；「复制 JSON」降级为高级折叠项。
- 新增 `examples/confirm_demo.py`：确认页文案与能力检测断言、自动定位（唯一/多份/未找到）、端到端 `locate → apply_contract dry-run → --write`、幂等与归档校验。

### Changed

- **SKILL.md 阶段 A**：A3 改为「提交后回对话说『好了』」；新增 A3.5「自动定位确认单」（`locate_confirm.py`，含多份/未找到时的处理规则）；§6 新手引导同步改为「助手自动归档」；§10 补确认单相关报错；§11 文件清单新增脚本；§12 补浏览器安全限制与回退说明。
- **红线补充**：不得要求用户移动/复制/重命名确认单文件，不得向普通用户解释 JSON 是什么；多份确认单时助手用最新一份并说明，或一句话请用户确认。
- `make demo` 现在运行四套端到端演示（对账 / L6 复核 / 结构映射 / 确认单定位）；CI 对应更新。
- 技能版本号 1.3.0（三处同步）。
- README 工作流与快速开始同步「确认单零文件操作」。

## [1.2.0] - 2026-11-03

新增两项能力：**L6 LLM 语义复核（交付门禁）** 与 **结构漂移自动识别 + 映射确认（阶段 0）**。技能仍保持 L1–L5 确定性校验不变，新增层只增加门禁与适配能力，不替代任何既有校验。

### Added

- **L6 LLM 语义复核（交付门禁）**
  - `make_review_packet.py`：生成 token 友好的复核包（关键数、逐 SKU / 贴花明细、文字断言、交叉引用、红线措辞扫描、合同条款、R1–R10 清单），取不到的数写 `null`+`warnings`，绝不臆造 0；除 `generated_at` 外确定性输出。
  - `apply_review.py`：校验复核报告（哈希绑定、R1–R10 全覆盖、结论一致性：阻断⇒不通过 / 有警告⇒不得通过），`--apply` 写入工作簿第 9 张表 `08-复核记录` 并生成 `_复核记录/复核锁.json`（pre/post sha256），`--check` 交付前校验「复核有效」；复核后任何保存都会使锁失效，防止「审旧交新」；不通过时只归档报告、不写表、不得交付；重复执行幂等。
  - 复核由**当前 harness 的 LLM** 执行（脚本不调用任何外部模型、不联网）；无文件写权限的 harness 可只输出报告 JSON 人工归档。
  - 新增 `references/复核协议.md`：R1–R10 逐项操作手册、复核包/报告 JSON schema、判定矩阵、08 表与复核锁留痕、外部工作簿复核回退。
  - 新增 `examples/review_demo.py`：8 表工作簿 → 复核包 → 08 表 → 复核锁 → 篡改失效 → 负例，约 4–5 秒。
- **结构漂移自动识别与映射确认（阶段 0）**
  - `inspect_schema.py`：只读探查主档/三份导出的表名、字段名、列顺序漂移，输出带置信度的映射建议（`proposed_mapping`）与 issue code（`SHEET_NAME_MISMATCH` / `FIELD_NAME_MISMATCH` / `FIELD_ORDER_DIFF` / `MULTI_CANDIDATE_SHEET` / `MISSING_FIELD` / `EXTRA_COLUMN`）。
  - `apply_mapping.py`：默认 dry-run 逐项核对，`--write` **必须**带 `--confirmed-by <确认人>`（用户确认门禁），写 `_数据映射.json`（schema `dealer-settlement/data-mapping@1`）并归档 `_确认记录/`；`--fix-master [--fix-order]` 可按规范修复主档表名/列名/列序/表序（先 dry-run、再确认、毫秒备份、原子替换），**永不修改三份源导出**。
  - 引擎消费映射：`load_master.py` 在客户根目录发现 `_数据映射.json` 时注入 `RECON_CFG['mapping']`；`recon_engine.py` 按映射选文件/工作表/表头行/列（显式优先于关键词启发），映射指向不存在的表/列时中文中止并指明是哪一项；**无映射时行为与 1.1.x 完全一致**（向后兼容）。
  - 新增 `references/数据映射与结构探查.md`：漂移类型、探查结果解读、用户确认话术、映射 JSON 结构、主档修复规则。
  - 新增 `examples/schema_demo.py`：制造 5 类漂移 → 自动识别 → 无映射中文中止 → 确认映射 → 带映射输出与基线逐单元格一致 → 主档修复+备份校验 → 负例，约 4 秒。

### Changed

- **工作流**：SKILL.md 新增「阶段 0 结构自检与映射确认（按需）」；阶段 C 拆为 C1 确定性校验 → C2 复核包 → C3 L6 复核（必过）→ C4 交付登记；§4 校验体系新增 L6 行，统一表述为「五层确定性校验（L1–L5）+ L6 LLM 语义复核」；§9 复核模式重写为「本次出账复核」与「外部工作簿复核」两种。
- **frontmatter**：description 增加「复核 / 验收 / 审计口径 / 字段名·表名不一致 / 列顺序不同 / 表头识别 / 映射确认」等触发词；compatibility 补充 L6 依赖（harness LLM + 工作簿读写）与无写权限回退。
- **文档**：SKILL.md 与 references 同步新能力（含 `_复核记录/`、`_数据映射.json`、`_确认记录/` 的落盘位置与回滚）。
- **README / Makefile / CI**：`make demo` 现在运行三套端到端演示（对账 / L6 复核 / 结构映射）；CI 对应步骤改为 `make demo`。
- 技能版本号 1.2.0（`plugin.json` / `.claude-plugin/plugin.json` / SKILL.md `metadata.version` 三处同步，校验器强制一致）。

### Fixed

- 文档与实现一致性校正：复核包 `figures.period_end`、`period` 为 `{start,end}` 对象（非 `"202610"`）、`--check` 可省略报告路径、`--report-dir` 同时决定报告与复核锁位置、`--apply` 不通过返回 1 而默认 dry-run 返回 0。
- 映射「哈希绑定」表述改为与实际一致：探查报告记录当次文件 sha256 供留痕，`apply_mapping` 不做强制哈希比对（换文件需重新探查）。

## [1.1.1] - 2026-11-03

补齐 v1.1.0 发布后发现的 CI 与 Claude Code 兼容层问题；技能内容与 v1.1.0 相同。

### Fixed

- **Claude Code 兼容层**：`.claude-plugin/marketplace.json` 的插件 `source` 由 `"."` 恢复为 `"./"`（`claude plugin validate` 只接受 `./` 开头的相对路径）；`.claude-plugin/plugin.json` 移除 `displayName`（当前 Claude Code 校验器报 `Unrecognized key`）。已在隔离 HOME 下实测 `claude plugin marketplace add` + `claude plugin install` 成功。
- **CI**：`actions/checkout`、`actions/setup-python`、`actions/upload-artifact` 升级到 v7（v4/v5 依赖已弃用的 Node 20）；Python 3.9 矩阵改为固定 `ubuntu-24.04`，因为 `ubuntu-latest`（26.04）没有 3.9 预编译包。
- **规范校验器**：新增「`SKILL.md` 的 `metadata.version` 必须与 `plugin.json` 的 `version` 一致」校验，避免发布版本漂移。

## [1.1.0] - 2026-11-03

本版本包含两项互相独立的改动：**打包层**从「单个 SKILL.md 仓库」升级为符合 Agent Plugins 1.0.0 的插件仓库；**脚本层**对 7 个 Python 脚本做了一轮审计式修复，并补齐可复现的演示与 CI。

### Added

- **Agent Plugins 1.0.0 打包**：新增根目录 `plugin.json`（封闭字段集、规范 `$schema`），技能迁入规范固定位置 `skills/dealer-settlement/`，任何实现 Agent Plugins 1.0.0 的客户端都可发现本技能（官方兼容清单：<https://agent-plugins.org/compatible-clients>）。
- **Claude Code 兼容层**：新增 `.claude-plugin/plugin.json` 与 `.claude-plugin/marketplace.json`，支持 `claude plugin marketplace add` 一键安装，也支持 `claude --plugin-dir` 本地加载。
- **单技能包构建**：`make package` 产出 `dist/dealer-settlement.zip`（技能目录位于 zip 根），保持 `unzip -d ~/.claude/skills/` 的传统安装方式可用。
- **可复现演示**：`examples/make_demo_data.py` 生成自洽的演示主档与三份月度导出；`examples/run_demo.py` 在临时目录端到端跑通并自检（8 表、恒等式、贴花复算、失败路径 stdout=0、确定性、无网络），CI 直接调用。
- **规范校验器**：`tools/validate.py` 校验插件清单、SKILL.md frontmatter、目录名匹配与相对引用（仅标准库、离线可用）。
- **CI**：GitHub Actions 在 Python 3.9 / 3.12 / 3.13 上执行规范校验、端到端演示与打包；新增 Dependabot 配置。
- **文档**：新增 `references/故障与回滚.md`（完整报错对照与回滚步骤，从 SKILL.md 下沉，保持渐进式披露）。
- 开源配套：`LICENSE`（MIT）、`CHANGELOG.md`、`CONTRIBUTING.md`、`.gitignore`、`.gitattributes`、`requirements.txt`、`Makefile`。

### Changed

- **README 重写**：安装矩阵覆盖「Agent Plugins 客户端 / Claude Code 插件 / 单技能包」三种形态，补充官方兼容客户端清单、规范符合性表与仓库结构说明。
- **SKILL.md 多 harness 适配**：新增「0. 运行约定」——用 `SKILL_DIR`（SKILL.md 所在目录）定位脚本、不假定 cwd；全部命令改为 `"$SKILL_DIR/scripts/..."`；给出 PowerShell 等价写法；移除 `present_files`、`AskUserQuestion` 等 harness 专有工具名，改为「能力 + 回退」写法。
- **frontmatter 合规**：改为 Agent Skills 规范字段（`name` / `description` / `license` / `compatibility` / `metadata`），删除非规范的 `agent_created`。
- **对账单条款文案参数化**：结算方式 / 结算基数 / 异议期 / 签回期改为从主档「合同」表读取；缺失时输出「合同约定的异议期限 / 签回期限」「主档未登记，以合同为准」，不再印死演示值。
- **references 一致性**：`客户主档设计.md` 由「五张表」更正为六表并补「贴花」表设计；区分「引擎硬性依赖 / 人工核对 / 脚本不读写」字段。

### Fixed

- **`recon_engine.py`（23 类）**：挂账 `.xlsx` 误用 `xlrd` 必崩 → 按扩展名分派并兼容首行标题行；源表脏值静默吞行 / 裸 traceback → 显式中止；工作表多候选静默取第一张 → 中止并列候选；无「次月销售」表或 `--asof`=期末时崩溃；恒等式 `assert` 在 `python -O` 下静默失效 → 显式自检；失败路径 stdout 非 0 字节（违反硬规矩）→ 输出移到保存成功之后；覆盖保护标记永远打印 `[新建]` → 现 `[覆盖→<文件名>]`；主体 / 日期 / 条款等演示文案硬编码并靠整簿 REPL 兜底（替换不到异议期 / 签回期 / 结算方式 / 结算基数）→ 全部按配置生成；盘盈被印成短缺；02 差异率除零；05 门店列在门店数≠5 时错位 → 动态列；06 结算属性按品名匹配 → 按条码分桶；07 文字断言与数据不符；主档路径缺失时静默跳过结算属性三态校验 → L1 中止；`RECON_CFG` 非法 JSON 被误报为「缺少必要信息」；销售日期仅支持 ISO → 多格式解析。
- **`load_master.py`（8 类）**：死条件分支；glob 通配符误匹配（文件名含 `[]?*`）；`base_dir` 相对路径随 CWD 漂移；`--period` 接受全角数字 / 非法年份；`--asof` / `--pay` 非法格式裸 traceback；合同条款字段未透传引擎；工作簿句柄未关闭。
- **`apply_contract.py`（9 类）**：重复 `import re`；死表达式；期间「贴花金额」漏检（L2）；状态枚举与文档不一致（补「已签回 / 已收款」）；新合同编号覆盖第 2 行旧合同 → 改为追加行；备份时间戳声称毫秒实为截断 1 位 → 补足毫秒；保存失败裸 traceback → 提示备份位置；文件 / 工作簿句柄未关闭。
- **`extract_contract.py`（5 类）**：docstring 未写 `.md` 支持；`.txt` / `.md` 仅 UTF-8，GBK 合同崩溃 → 编码兜底；损坏 `.docx`、文件不存在友好中止；PDF 全局状态跨调用污染 `needs_ocr`；输出文件句柄 / 目录不存在。
- **`make_digest.py` / `make_confirm_html.py` / `make_master.py`（6 类）**：主档 / 候选 JSON 缺失或损坏时裸 traceback；输出目录不存在；工作簿与文件句柄未关闭；`make_master.py` docstring 与实际行为不符；重复 import。
- **文档一致性**：六表表述、恒等式条数与 ①–⑦ 结转编号区分、确认字段数量（22 项）、红线编号 1–16 连续、8 张输出表名与 `create_sheet` 实测一致、断言符号、期间表「脚本自动追加」等错误表述。

### 已知限制

- 真实 `.xls` 挂账样本尚未在真机实测（代码已按扩展名分派，`xlrd` 懒加载）；`.xlsx` / `.xlsm` 路径已在演示中覆盖。
- 同一客户多份合同时，不带主键的合同字段仍写入第一份合同（合同编号已改为追加新行）；续签合同请走确认页人工确认。
- 引擎对「多张同名候选工作表」改为严格中止（与文档一致），客户导出若混入辅助表需先规范表名。
