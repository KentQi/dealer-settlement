# 贡献指南

感谢参与 dealer-settlement。本仓库既是 **Agent Skill**，也是 **Agent Plugin**（Agent Plugins 1.0.0），
任何改动都要同时满足两份规范。

## 环境准备

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
make check   # 清单/规范校验 + 语法编译
make demo    # 生成演示数据并端到端跑通对账流程
```

## 目录约定

| 路径 | 说明 |
|---|---|
| `skills/dealer-settlement/` | 技能本体；目录名必须与 SKILL.md 的 `name` 一致 |
| `skills/dealer-settlement/scripts/` | 技能脚本（技能运行时通过 SKILL_DIR 定位） |
| `skills/dealer-settlement/references/` | 按需加载的参考文档 |
| `skills/dealer-settlement/assets/` | 模板等静态资源 |
| `examples/` | 演示数据生成器与端到端演示（CI 使用） |
| `tools/` | 仓库校验脚本（不随插件分发） |

## 提交前检查

1. `make check` 与 `make demo` 必须全绿。
2. 修改 `SKILL.md` 时保持 frontmatter 合规：`name` 与目录同名、`description` ≤1024 字符；
   harness 专有能力一律写成「能力 + 回退」，不要硬编码某一家 harness 的工具名。
3. 版本号三处同步：`plugin.json`、`.claude-plugin/plugin.json`、`SKILL.md` 的 `metadata.version`。
4. 更新 `CHANGELOG.md`（Keep a Changelog 格式）。

## 红线

- **不提交真实客户数据**：主体、税号、账号、价格一律用演示/虚构值。
- **不引入网络调用**：对账必须完全本地执行，不上传任何业务数据。
- **不弱化五层校验**：缺失即阻断、不静默兜底是设计前提，相关行为必须保留。
- **不静默改变 CLI 与输出文件名**：SKILL.md 已文档化，破坏性变更需同步文档并升级次版本号。
- 报错文案保持中文，并给出可整行复制的修复命令。
