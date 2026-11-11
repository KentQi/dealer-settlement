# dealer-settlement — 铺货代销对账 Skill / Agent Plugin

[![CI](https://github.com/KentQi/dealer-settlement/actions/workflows/ci.yml/badge.svg)](https://github.com/KentQi/dealer-settlement/actions/workflows/ci.yml)
[![Agent Plugins 1.0.0](https://img.shields.io/badge/Agent%20Plugins-1.0.0-informational)](https://agent-plugins.org/specification)
[![Agent Skills](https://img.shields.io/badge/Agent%20Skills-specification-informational)](https://agentskills.io/specification)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)

面向 AI Agent 的「铺货代销对账」作业包：从经销合同建档，到月度对账出 10 表工作簿（对账单 / 库存差异 / 贴花 / 批次效期预警 / 试用装与配赠），全程**五层确定性校验（L1–L5）+ L6 LLM 语义复核**，缺失即阻断；上游还能自动识别客户导出/主档的**表名、字段名、字段顺序漂移**，经用户确认后落映射继续跑。

本仓库同时是：

- **Agent Plugin**（根目录 `plugin.json` + `skills/`，遵循 [Agent Plugins 1.0.0](https://agent-plugins.org/specification)）——可移植包，插件客户端自动发现技能
- **Agent Skill**（`skills/dealer-settlement/`，遵循 [Agent Skills](https://agentskills.io/specification)）——也可单独打包，放进只认 `SKILL.md` 的 harness 的 skills 根目录
- **Claude Code 插件**（`.claude-plugin/` 兼容层）——支持 `claude plugin marketplace add` / `--plugin-dir`

> 技能规程见 [skills/dealer-settlement/SKILL.md](skills/dealer-settlement/SKILL.md)。仓库内所有主体、品名、条码、价格均为**演示（虚构）数据**。

## 它做什么

```
【一次性建档】
合同原件 ──抽取──▶ 候选字段(带原文证据) ──HTML确认页(人工逐项确认)──▶ 助手自动定位确认单 ──▶ 写入客户主档（六表）

【数据接入（按需）】
主档/三份导出 ──inspect_schema──▶ 漂移报告+映射建议 ──用户逐项确认──▶ _数据映射.json ──▶ 引擎按映射取数

【每月对账】
三份月度导出(挂账/销售/库存) + 主档 ──▶ 10 表对账工作簿 + 贴花 + 效期预警 + 促销物料 + 自检
                                              │
                                              ▼
                        L6 LLM 语义复核（R1–R10 清单）──▶ 10-复核记录 + 复核锁 ──▶ 交付
```

- **确认单零文件操作**：确认页在浏览器里逐项确认后，用户只需回对话说一句「好了」；助手用 `locate_confirm.py` 自动从浏览器下载目录（或客户目录——即存放该客户资料的目录）找到确认单并归档，**不需要用户搬运/重命名 JSON**；支持 File System Access API 的浏览器，页面上还有「存到客户目录…」按钮可直接保存
- **对账口径**：铺货代销、按实际动销结算；六条恒等式 + 贴花计算规则，含时点回加
- **贴花模块**：按「结算类动销数量 × 当期生效标准」计算，SCD2 生效区间，脏值 / 区间重叠 / 缺失一律阻断
- **批次效期预警（附表 08）**：读挂账表的批次号与有效期，按「期初批次 − 本期动销」的 FIFO 扣减算期末结存，按已过期 / 30 / 90 / 180 天分档；批次结存与账面勾稽、后进先出迹象单列。无批次数据时只出说明页，不推测
- **试用装 / 中小样 / 配赠（附表 09）**：从挂账表的「物料属性」分流，不参与货款结算；乙方提供使用数据则列示差异，未提供则按「已出库即计列（未经确认）」列示，不编造使用量
- **五层确定性校验**：入口必填 → 写库类型 → 计算自检 → 结构守卫 → 输出防覆盖；失败时 stderr 给中文原因与可复制的修复命令，stdout 保持 0 字节
- **L6 LLM 语义复核（交付门禁）**：脚本生成 token 友好的复核包（关键数 / 逐 SKU / 文字断言 / 交叉引用 / 红线扫描），由当前 harness 的 LLM 按 R1–R10 独立复算并判定；结论写入工作簿 `10-复核记录`，并用「复核锁」防止**审旧交新**——不通过不得交付，复核后再改文件即失效、必须重评
- **结构漂移自动识别与映射确认**：表名 / 字段名 / 字段顺序不一致时，自动探查并给出带置信度的映射建议；**低置信项必须由用户逐项裁决**，确认后写入 `_数据映射.json` 供引擎确定性取数；主档可安全修复（先 dry-run、确认、再备份写入），永不修改客户源导出
- **LLM 不联网、不调用外部模型**：复核与识别都由当前 harness 的 LLM 完成，脚本只提供结构化底稿、校验规则与留痕

## 兼容的 harness

Agent Plugins 1.0.0 的可移植核只定义 `plugin.json` + `skills/`（+ 可选 `mcp.json`），本仓库不绑定特定客户端：只要客户端按该规范加载插件，就能把本仓库装进来。官方兼容清单目前列出：

VS Code · Cursor · GitHub Copilot · ChatGPT & Codex · Kiro · Hermes Agent · OpenClaw · Grok Bot · NanoClaw · OpenHands · Perplexity Computer

（完整清单与各客户端支持的组件：<https://agent-plugins.org/compatible-clients>）

按客户端类型分四条路：

| Harness | 走哪条路 |
|---|---|
| 上述 Agent Plugins 客户端 | 把仓库根目录（`plugin.json` + `skills/`）作为插件包加入其插件目录或市场 |
| Claude Code | 作为插件安装，或本地 `--plugin-dir` 直接加载 |
| 仅支持 Agent Skills 的 harness | 解压技能包到该 harness 的 skills 根目录 |
| WorkBuddy | 解压技能包到 `~/.workbuddy/skills/` |

具体命令见下节。技能内所有命令都用 `SKILL_DIR`（SKILL.md 所在目录）定位脚本，不依赖当前工作目录；同时给出 bash 与 PowerShell 写法（见 SKILL.md 第 0 节）。

## 安装与快速开始

### 作为插件安装（推荐）

```bash
# Claude Code：装成插件
claude plugin marketplace add KentQi/dealer-settlement
claude plugin install dealer-settlement@dealer-settlement

# Claude Code：不装市场，直接用本地仓库
claude --plugin-dir /path/to/dealer-settlement
```

其他 Agent Plugins 客户端把仓库根目录当作插件包，加到各自的插件目录或市场即可（见上表）。

### 作为单技能包安装

```bash
make package    # 生成 dist/dealer-settlement.zip（技能包）+ dist/dealer-settlement-<版本>-source.zip（源码包）
                # 没有 make 时等价命令：python3 tools/package_zip.py
unzip dist/dealer-settlement.zip -d ~/.claude/skills/

# 或直接软链（目录名必须保持 dealer-settlement）
ln -s "$PWD/skills/dealer-settlement" ~/.claude/skills/dealer-settlement
```

装到别的 harness 时把 `~/.claude/skills/` 换成该 harness 的 skills 根目录（WorkBuddy 为 `~/.workbuddy/skills/`）。

### 跑一遍可复现的演示

```bash
pip install -r requirements.txt
make demo      # 四套端到端演示：对账主流程 / L6 复核门禁 / 结构漂移识别 / 确认单自动定位
# 也可单独跑：
python3 examples/run_demo.py      # 生成演示数据 → 10 表对账 → 自检（含效期/促销两种分支）
python3 examples/review_demo.py   # 复核包 → 复核记录 → 复核锁 → 防审旧交新
python3 examples/schema_demo.py   # 制造表名/列名/顺序漂移 → 自动识别 → 确认映射 → 重跑通过
python3 examples/confirm_demo.py  # 确认页文案 → 自动定位确认单 → 归档写入主档
```

演示数据全部生成在临时目录，不会写入任何业务目录。真实业务使用见 SKILL.md「3. 工作流」。

依赖：`python3` 3.9+，以及 `requirements.txt` 里的三个库（openpyxl / xlrd / pdfplumber）。技能脚本纯本地执行，不访问网络。

## 仓库结构

```text
.
├── plugin.json                     # Agent Plugins 1.0.0 可移植清单（封闭字段集）
├── skills/dealer-settlement/       # 技能本体（可移植内容）
│   ├── SKILL.md                    # frontmatter + 作业规程
│   ├── scripts/                    # 建档、对账、复核（L6）、结构映射、校验脚本
│   ├── references/                 # 主档设计 / 复核协议 / 数据映射 / 故障回滚（按需加载）
│   └── assets/                     # 主档模板
├── examples/                       # 演示数据生成器 + 四套端到端演示（CI 使用）
├── tools/                          # 仓库工具：规范校验 + 打包（不随技能分发）
├── .claude-plugin/                 # Claude Code 原生插件与 marketplace 清单（兼容层）
├── .github/workflows/ci.yml        # 多版本 Python CI
├── Makefile                        # make check / demo / package
├── CHANGELOG.md / CONTRIBUTING.md
└── requirements.txt
```

## 规范符合性

| 规范 | 要求 | 本仓库 |
|---|---|---|
| Agent Plugins 1.0.0 | 根 `plugin.json`、封闭字段、`$schema` 固定值、name 规则 | ✅ 符合 |
| Agent Plugins 1.0.0 | 技能位于固定位置 `skills/<name>/SKILL.md`（不递归发现） | ✅ `skills/dealer-settlement/` |
| Agent Skills | frontmatter `name` 与目录同名、`description` ≤1024、规范字段集 | ✅ 见 [SKILL.md](skills/dealer-settlement/SKILL.md) |
| Agent Plugins 1.0.0 | `mcp.json` 可选；无 MCP 服务器时缺失不算错误 | ✅ 未提供（本技能不需要） |

`make check`（即 `python3 tools/validate.py` + 语法编译）会机器校验上述要求；CI 在 Python 3.9 / 3.12 / 3.13 上执行。

## 开发

| 命令 | 作用 |
|---|---|
| `make check` | 校验 `plugin.json`、`SKILL.md` 是否符合 Agent Plugins / Agent Skills 规范，并做语法编译 |
| `make demo` | 跑四套端到端演示（见上「跑一遍可复现的演示」） |
| `make package` | 构建两个包：`dist/dealer-settlement.zip`（技能安装包）、`dist/dealer-settlement-<版本>-source.zip`（完整源码）；文件名统一 UTF-8，跨平台解压不乱码 |
| `make clean` | 清理构建产物与 `__pycache__` |

贡献前请阅读 [CONTRIBUTING.md](CONTRIBUTING.md)。

## 适配边界

仅适用「铺货代销 + 按实际动销结算 + 月末对账」；买断、按发货结算、退货冲减、返利/折让计价不在恒等式覆盖内。单据版式按化妆品类经销商的常见单据形态打磨，换行业需复核版式。

## 数据与安全

- 技能全程本地运行，不上传任何业务数据；导出、主档、确认记录都留在用户自己的目录。
- 仓库内**不含任何真实客户、合同、价格或对账数据**：演示主体、品名、条码、价格全部虚构，业务场景描述已做去标识化处理，无法反推到具体企业。**请勿提交真实业务数据**（见 CONTRIBUTING 红线）。
- 对账结果不构成审计意见；L6 LLM 复核也不构成审计/会计师意见，差异归因、税率、赔偿条款均需人工与双方书面确认。
- 复核包、复核报告、映射确认记录都落在客户目录（`_复核记录/`、`_确认记录/`），便于事后追溯；不随任何网络请求外发。

## 许可

[MIT](LICENSE) © 2026 KentQi
