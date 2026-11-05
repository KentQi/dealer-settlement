#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""打包发布产物。

产出两个文件到 dist/：
  1. dealer-settlement.zip            —— 技能安装包（只含 skills/dealer-settlement/，扁平结构）
  2. dealer-settlement-<ver>-source.zip —— 完整源码包（仓库全部文件，带版本前缀目录）

为什么不用系统 `zip`：macOS 自带的 zip 会把中文文件名按本地编码写入且不设 UTF-8
标志（flag 0x800），Windows 解压后文件名变成乱码。这里统一用 Python zipfile，
并在写入时显式设置 UTF-8 标志。
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
    info = zipfile.ZipInfo.from_file(src, arcname)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.flag_bits |= 0x800          # UTF-8 文件名标志
    with open(src, "rb") as fh, zf.open(info, "w") as out:
        out.write(fh.read())


def build_skill() -> Path:
    src = ROOT / "skills" / "dealer-settlement"
    if not src.is_dir():
        raise SystemExit(f"[打包失败] 找不到技能目录：{src}")
    out = DIST / "dealer-settlement.zip"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
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


def _report(path: Path) -> None:
    with zipfile.ZipFile(path) as zf:
        infos = zf.infolist()
        bad = [i.filename for i in infos
               if any(ord(ch) > 127 for ch in i.filename) and not (i.flag_bits & 0x800)]
    print(f"  {path.name}: {len(infos)} 个文件, {path.stat().st_size / 1024:.1f} KB, "
          f"文件名编码异常 {len(bad)}")
    if bad:
        raise SystemExit("[打包失败] 以下文件名未标记 UTF-8：" + repr(bad[:5]))


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
    skill = build_skill()
    source = build_source(version)
    _report(skill)
    _report(source)
    print(f"输出目录：{DIST}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
