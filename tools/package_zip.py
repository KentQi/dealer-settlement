#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""打包发布产物。

产出两个文件到 dist/：
  1. dealer-settlement.zip            —— 技能安装包（只含 skills/dealer-settlement/，扁平结构）
  2. dealer-settlement-<ver>-source.zip —— 完整源码包（仓库全部文件，带版本前缀目录）

为什么不用系统 `zip`：macOS 自带的 zip 会把中文文件名按本地编码写入且不设 UTF-8
标志（flag 0x800），Windows 解压后文件名变成乱码。这里统一用 Python zipfile（它会为
非 ASCII 文件名写入 UTF-8 标志），并在打包后逐条断言，不通过即失败。
"""
import json
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DIST = ROOT / "dist"

EXCLUDE_DIRS = {".git", "dist", "__pycache__", ".pytest_cache", ".idea", ".vscode", ".venv"}
EXCLUDE_SUFFIX = {".pyc", ".pyo"}
EXCLUDE_NAMES = {".DS_Store", ".commitmsg"}


def _iter_files(base: Path):
    for path in sorted(base.rglob("*")):
        if not path.is_file():
            continue
        rel = path.relative_to(base)
        if set(rel.parts) & EXCLUDE_DIRS:
            continue
        if path.name in EXCLUDE_NAMES or path.suffix in EXCLUDE_SUFFIX:
            continue
        if any(part.endswith(".bak") for part in rel.parts):
            continue
        yield path, rel


def _write(zf: zipfile.ZipFile, src: Path, arcname: str) -> None:
    # 非 ASCII 文件名的 UTF-8 标志（0x800）由 zipfile 自动写入；成品由 _report() 断言
    info = zipfile.ZipInfo.from_file(src, arcname)
    info.compress_type = zipfile.ZIP_DEFLATED
    with open(src, "rb") as fh, zf.open(info, "w") as out:
        out.write(fh.read())


def build_skill() -> Path:
    src = ROOT / "skills" / "dealer-settlement"
    if not src.is_dir():
        raise SystemExit(f"[打包失败] 找不到技能目录：{src}")
    lic = ROOT / "LICENSE"
    if not lic.is_file():
        raise SystemExit("[打包失败] 缺少 LICENSE：MIT 要求随分发副本携带许可声明")
    out = DIST / "dealer-settlement.zip"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        _write(zf, lic, "dealer-settlement/LICENSE")
        for path, rel in _iter_files(src):
            _write(zf, path, f"dealer-settlement/{rel.as_posix()}")
    return out


def build_source(version: str) -> Path:
    out = DIST / f"dealer-settlement-{version}-source.zip"
    prefix = f"dealer-settlement-{version}"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for path, rel in _iter_files(ROOT):
            _write(zf, path, f"{prefix}/{rel.as_posix()}")
    return out


def _report(path: Path, required=()) -> None:
    with zipfile.ZipFile(path) as zf:
        infos = zf.infolist()
        names = {i.filename for i in infos}
        bad = [i.filename for i in infos
               if any(ord(ch) > 127 for ch in i.filename) and not (i.flag_bits & 0x800)]
        missing = [r for r in required if r not in names]
    print(f"  {path.name}: {len(infos)} 个文件, {path.stat().st_size / 1024:.1f} KB, "
          f"文件名编码异常 {len(bad)}")
    if bad:
        raise SystemExit("[打包失败] 以下文件名未标记 UTF-8：" + repr(bad[:5]))
    if missing:
        raise SystemExit("[打包失败] 包内缺少必要文件：" + repr(missing))


def main() -> int:
    try:
        version = json.loads((ROOT / "plugin.json").read_text(encoding="utf-8"))["version"]
    except Exception as exc:
        raise SystemExit(f"[打包失败] 读取 plugin.json 失败：{exc}")

    if not DIST.is_dir():
        DIST.mkdir(parents=True)
    for old in DIST.glob("*.zip"):
        old.unlink()

    print(f"打包 dealer-settlement {version}")
    prefix = f"dealer-settlement-{version}"
    skill = build_skill()
    source = build_source(version)
    _report(skill, required=("dealer-settlement/LICENSE", "dealer-settlement/SKILL.md"))
    _report(source, required=(f"{prefix}/LICENSE", f"{prefix}/plugin.json"))
    print(f"输出目录：{DIST}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
