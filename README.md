# dealer-settlement — 铺货代销对账 Skill

面向 AI Agent 的 **Agent Skill**（遵循 Anthropic SKILL.md 规范）：从经销合同建档，到月度对账出 8 表工作簿（对账单 / 库存差异 / 贴花费用），全程五层校验、缺失即阻断。

> 详细规程见 [SKILL.md](SKILL.md)。本仓库内所有主体、品名、条码、价格均为**演示数据（虚构）**。

## 它做什么

```
【一次性建档】
合同原件 ──抽取──▶ 候选字段(带原文证据) ──HTML确认页(人工逐项确认)──▶ 写入客户主档（六表）

【每月对账】
三份月度导出(挂账/销售/库存) + 主档 ──▶ 8 表对账工作簿 + 贴花费用计算 + 自检报告
```

- **对账口径**：铺货代销、按实际动销结算；七条恒等式含时点回加与贴花
- **贴花模块**：按「结算类动销数量 × 当期生效标准」计算，SCD2 生效区间，脏值/区间重叠/缺失一律阻断
- **五层校验**：入口必填 → 写库类型 → 计算自检 → 结构守卫 → 输出防覆盖；失败时 stderr 给中文原因与可复制的修复命令，stdout 保持 0 字节

## 安装（任选其一）

| Harness | 安装位置 |
|---|---|
| Claude Code / Claude Desktop | `~/.claude/skills/dealer-settlement/` |
| WorkBuddy | `~/.workbuddy/skills/dealer-settlement/` |
| 其他兼容 SKILL.md 规范的 Agent | 其 skills 根目录下解压同名文件夹即可 |

```bash
unzip dealer-settlement.zip -d ~/.claude/skills/   # 目录名必须保持 dealer-settlement
```

依赖：`python3` + `pip install openpyxl xlrd pdfplumber`（xlrd ≥2.0 仅读 .xls；写 demo .xls 需 `xlwt`）。

## 快速开始（demo 数据）

```bash
python3 scripts/make_master.py /tmp/demo/客户/_主档.xlsx     # 生成演示主档（六表）
# 准备三份导出后：
export RECON_CFG=$(python3 scripts/load_master.py "/tmp/demo/客户/_主档.xlsx" --period 202610 --asof 2026-11-02)
python3 scripts/recon_engine.py                              # 生成 8 表对账工作簿
```

## 适配边界

仅适用「铺货代销 + 按实际动销结算 + 月末对账」；买断/按发货结算/退货冲减/返利计价不在恒等式覆盖内。单据版式按某化妆品品牌经销商场景打磨，换行业需复核版式。
