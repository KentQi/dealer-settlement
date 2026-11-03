#!/usr/bin/env python3
"""校验本仓库是否符合 Agent Plugins 1.0.0 与 Agent Skills 规范。

只依赖 Python 标准库，CI 与本地 make check 都会调用。
校验项：
  1. plugin.json          —— 封闭字段集、$schema、name 规则与字段类型
  2. .claude-plugin/*     —— JSON 合法且与根清单的名称/版本一致（Claude Code 兼容层）
  3. skills/*/SKILL.md    —— frontmatter 规则、name 与父目录同名、description 长度
  4. SKILL.md 内的一级相对链接是否指向真实文件
  5. mcp.json（若存在）    —— 顶层字段与服务器条目结构

用法：
    python3 tools/validate.py
退出码：0 = 通过；1 = 存在错误。
"""
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PLUGIN_SCHEMA = "https://agent-plugins.org/schemas/1.0.0/plugin.schema.json"
MCP_SCHEMA = "https://agent-plugins.org/schemas/1.0.0/mcp.schema.json"

PLUGIN_KEYS = {"$schema", "name", "version", "description", "author",
               "homepage", "repository", "license", "keywords", "extensions"}
AUTHOR_KEYS = {"name", "email", "url"}
SKILL_KEYS = {"name", "description", "license", "compatibility", "metadata", "allowed-tools"}

# Agent Plugins 规范 §5.5
PLUGIN_NAME_RE = re.compile(r"^(?!.*(?:--|\.\.))[a-z0-9](?:[a-z0-9.-]*[a-z0-9])?$")
# Agent Skills 规范：小写字母/数字/连字符，首尾非连字符，无连续连字符
SKILL_NAME_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
LINK_RE = re.compile(r"\[[^\]]*\]\(([^)\s]+)\)")

errors = []
warnings = []
checks = 0


def ok(msg):
    global checks
    checks += 1
    print("  \u2713 " + msg)


def err(msg):
    errors.append(msg)
    print("  \u2717 " + msg)


def warn(msg):
    warnings.append(msg)
    print("  ! " + msg)


def rel(path):
    return os.path.relpath(path, ROOT)


def load_json(path, label):
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except FileNotFoundError:
        err(label + "：文件不存在")
    except json.JSONDecodeError as exc:
        err(label + "：不是合法 JSON（" + str(exc) + "）")
    return None


def check_plugin_manifest():
    print("plugin.json")
    data = load_json(os.path.join(ROOT, "plugin.json"), "plugin.json")
    if data is None:
        return None
    if not isinstance(data, dict):
        err("plugin.json：顶层必须是 JSON 对象")
        return None

    unknown = sorted(set(data) - PLUGIN_KEYS)
    if unknown:
        err("plugin.json：不允许的顶层字段 " + "、".join(unknown))

    if data.get("$schema") != PLUGIN_SCHEMA:
        err("plugin.json：$schema 必须为 " + PLUGIN_SCHEMA)

    name = data.get("name")
    if not isinstance(name, str) or not name:
        err("plugin.json：缺少必填字段 name（非空字符串）")
    elif not (1 <= len(name) <= 64) or not PLUGIN_NAME_RE.match(name):
        err("plugin.json：name 不符合规则（1-64 字符，小写字母/数字/连字符/句点，"
            "首尾为字母数字，禁止连续 -- 或 ..）：" + repr(name))

    for key in ("version", "description", "homepage", "repository", "license"):
        if key in data and not isinstance(data[key], str):
            err("plugin.json：" + key + " 必须是字符串")
    if "keywords" in data and (not isinstance(data["keywords"], list)
                               or any(not isinstance(k, str) for k in data["keywords"])):
        err("plugin.json：keywords 必须是字符串数组")
    if "author" in data:
        author = data["author"]
        if not isinstance(author, dict):
            err("plugin.json：author 必须是对象")
        else:
            extra = sorted(set(author) - AUTHOR_KEYS)
            if extra:
                err("plugin.json：author 只允许 name/email/url，发现 " + "、".join(extra))
            for key in author:
                if not isinstance(author[key], str):
                    err("plugin.json：author." + key + " 必须是字符串")
    if "extensions" in data and not isinstance(data["extensions"], dict):
        warn("plugin.json：extensions 不是对象，客户端会报告并忽略（非致命）")

    ok("字段封闭性 / $schema / name / 类型")
    return data


def check_skill(skill_dir, plugin):
    name = os.path.basename(skill_dir)
    md_path = os.path.join(skill_dir, "SKILL.md")
    print("skills/" + name + "/SKILL.md")
    try:
        with open(md_path, encoding="utf-8") as fh:
            text = fh.read()
    except OSError:
        err("无法读取 " + rel(md_path))
        return

    front, body = parse_frontmatter(text)
    if front is None:
        err(rel(md_path) + "：缺少合法 YAML frontmatter（首行 --- 与结束 ---）")
        return

    unknown = sorted(set(front) - SKILL_KEYS)
    if unknown:
        warn(rel(md_path) + "：非规范 frontmatter 字段 " + "、".join(unknown)
             + "（建议移入 metadata）")

    fm_name = front.get("name")
    if not isinstance(fm_name, str) or not fm_name:
        err("frontmatter：缺少必填字段 name")
    else:
        if len(fm_name) > 64 or not SKILL_NAME_RE.match(fm_name):
            err("frontmatter：name 不符合规则（≤64，小写字母/数字/连字符，首尾非连字符）："
                + repr(fm_name))
        if fm_name != name:
            err("frontmatter：name 必须与父目录名相同（" + name + "）：" + repr(fm_name))

    desc = front.get("description")
    if not isinstance(desc, str) or not desc.strip():
        err("frontmatter：缺少必填字段 description")
    elif len(desc) > 1024:
        err("frontmatter：description 超长（" + str(len(desc)) + " > 1024 字符）")

    compat = front.get("compatibility")
    if isinstance(compat, str) and len(compat) > 500:
        err("frontmatter：compatibility 超长（> 500 字符）")
    if "metadata" in front:
        meta = front["metadata"]
        if not isinstance(meta, dict):
            err("frontmatter：metadata 必须是键值映射")
        else:
            for key, value in meta.items():
                if not isinstance(value, str):
                    err("frontmatter：metadata." + str(key) + " 的值必须是字符串")
            ver = meta.get("version")
            if (plugin is not None and isinstance(ver, str) and plugin.get("version")
                    and ver != plugin["version"]):
                err("frontmatter：metadata.version（" + ver + "）与 plugin.json version（"
                    + str(plugin["version"]) + "）不一致")

    ok("frontmatter 字段与命名规则")

    missing = []
    for target in LINK_RE.findall(body):
        if "://" in target or target.startswith(("#", "/", "mailto:")):
            continue
        target = target.split("#", 1)[0].split("?", 1)[0]
        if not target.startswith(("scripts/", "references/", "assets/", "./")):
            continue
        if not os.path.isfile(os.path.join(skill_dir, target)):
            missing.append(target)
    if missing:
        err("SKILL.md 引用了不存在的文件：" + "、".join(sorted(set(missing))))
    else:
        ok("SKILL.md 引用文件均存在")

    body_lines = len(body.splitlines())
    if body_lines > 500:
        warn("SKILL.md 正文 " + str(body_lines) + " 行，超过建议的 500 行")


def parse_frontmatter(text):
    """最小 YAML 子集解析：顶层键 + 单层映射（metadata）。返回 (dict|None, body)。"""
    lines = text.splitlines()
    if not lines or lines[0].strip() != "---":
        return None, text
    end = None
    for idx in range(1, len(lines)):
        if lines[idx].strip() == "---":
            end = idx
            break
    if end is None:
        return None, text

    front = {}
    current = None
    for raw_line in lines[1:end]:
        if not raw_line.strip() or raw_line.lstrip().startswith("#"):
            continue
        if raw_line[0] in " \t":
            if current is None:
                return None, text
            key, sep, value = raw_line.strip().partition(":")
            if not sep:
                return None, text
            if not isinstance(front.get(current), dict):
                front[current] = {}
            front[current][key.strip()] = unquote(value.strip())
            continue
        key, sep, value = raw_line.partition(":")
        if not sep:
            return None, text
        current = key.strip()
        value = value.strip()
        front[current] = unquote(value) if value else {}
    return front, "\n".join(lines[end + 1:])


def unquote(value):
    if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
        return value[1:-1]
    return value


def check_claude_plugin(plugin):
    print(".claude-plugin/（Claude Code 兼容层）")
    cc = load_json(os.path.join(ROOT, ".claude-plugin", "plugin.json"),
                   ".claude-plugin/plugin.json")
    if cc is not None:
        if plugin is not None:
            for key in ("name", "version", "license"):
                if key in plugin and cc.get(key) != plugin.get(key):
                    err(".claude-plugin/plugin.json：" + key + " 与根 plugin.json 不一致（"
                        + repr(cc.get(key)) + " != " + repr(plugin.get(key)) + "）")
        if not cc.get("name"):
            err(".claude-plugin/plugin.json：缺少 name")
        ok("清单字段与根清单一致")

    market = load_json(os.path.join(ROOT, ".claude-plugin", "marketplace.json"),
                       ".claude-plugin/marketplace.json")
    if market is not None:
        if not market.get("name"):
            err(".claude-plugin/marketplace.json：缺少 name")
        if not isinstance(market.get("owner"), dict) or not market["owner"].get("name"):
            err(".claude-plugin/marketplace.json：缺少 owner.name")
        entries = market.get("plugins")
        if not isinstance(entries, list) or not entries:
            err(".claude-plugin/marketplace.json：plugins 必须是非空数组")
        else:
            for entry in entries:
                if not isinstance(entry, dict) or not entry.get("name") or not entry.get("source"):
                    err(".claude-plugin/marketplace.json：每个插件条目需要 name 与 source")
                    continue
                if plugin is not None and entry.get("name") != plugin.get("name"):
                    warn("marketplace 插件条目名与根清单 name 不一致：" + repr(entry.get("name")))
                source = entry.get("source")
                # Claude Code 要求相对路径插件源以 "./" 开头（"." 会被 claude plugin validate 拒绝）
                if isinstance(source, str) and not source.startswith("./"):
                    err(".claude-plugin/marketplace.json：" + repr(entry.get("name"))
                        + " 的 source 必须以 './' 开头，当前 " + repr(source))
        ok("marketplace.json 结构")


def check_mcp():
    path = os.path.join(ROOT, "mcp.json")
    if not os.path.exists(path):
        return
    print("mcp.json")
    data = load_json(path, "mcp.json")
    if data is None:
        return
    if set(data) != {"$schema", "mcpServers"}:
        err("mcp.json：顶层字段必须且只能是 $schema 与 mcpServers")
    if data.get("$schema") != MCP_SCHEMA:
        err("mcp.json：$schema 必须为 " + MCP_SCHEMA)
    servers = data.get("mcpServers")
    if not isinstance(servers, dict):
        err("mcp.json：mcpServers 必须是对象")
        return
    for name, server in servers.items():
        if not isinstance(server, dict):
            err("mcp.json：" + name + " 必须是对象")
            continue
        kind = server.get("type")
        if kind == "stdio":
            command = server.get("command")
            if not isinstance(command, str) or not command:
                err("mcp.json：" + name + " 缺少 command")
            env = server.get("env") or {}
            if isinstance(env, dict) and {"PLUGIN_ROOT", "PLUGIN_DATA"} & set(env):
                err("mcp.json：" + name + " 的 env 不得包含保留变量 PLUGIN_ROOT/PLUGIN_DATA")
        elif kind in ("streamable-http", "sse"):
            url = server.get("url") or ""
            if not isinstance(url, str) or not url.startswith(("http://", "https://")):
                err("mcp.json：" + name + " 的 url 必须是绝对 HTTP(S) URL")
        else:
            err("mcp.json：" + name + " 的 type 必须是 stdio / streamable-http / sse")
    ok("服务器条目结构")


def main():
    print("dealer-settlement · Agent Plugins 1.0.0 / Agent Skills 规范校验\n")
    plugin = check_plugin_manifest()

    skills_root = os.path.join(ROOT, "skills")
    if not os.path.isdir(skills_root):
        err("缺少 skills/ 目录（Agent Plugins §6.1 固定组件位置）")
    else:
        found = 0
        for entry in sorted(os.listdir(skills_root)):
            skill_dir = os.path.join(skills_root, entry)
            if not os.path.isdir(skill_dir):
                continue
            if not os.path.isfile(os.path.join(skill_dir, "SKILL.md")):
                warn("skills/" + entry + "：缺少 SKILL.md，客户端会跳过该目录")
                continue
            found += 1
            check_skill(skill_dir, plugin)
        if found == 0:
            err("skills/ 下没有发现任何技能（每个技能目录必须含 SKILL.md）")

    check_claude_plugin(plugin)
    check_mcp()

    print()
    if errors:
        print("失败：" + str(len(errors)) + " 个错误，" + str(len(warnings)) + " 个警告")
        for msg in errors:
            print("  - " + msg)
        return 1
    print("通过：" + str(checks) + " 项检查，" + str(len(warnings)) + " 个警告")
    return 0


if __name__ == "__main__":
    sys.exit(main())
