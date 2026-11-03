# Changelog

本项目的所有重要变更记录于此，格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，
版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

## [Unreleased]

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
