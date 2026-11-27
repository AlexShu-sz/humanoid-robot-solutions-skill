#!/usr/bin/env python3
"""从同一次提交生成两个平台的技能安装包。

仓库布局：
  skill/                共用的 SKILL.md、三份规则、通用脚本
  platforms/<平台>/     平台执行说明与确实依赖该平台的脚本
  templates/            PPT / Excel 模板，各保留一份

产物：dist/industrial-robot-solution_v<版本>_<平台>.zip
  包内根目录固定为 industrial-robot-solution/，SKILL.md 在包根，
  模板从 templates/ 复制到 assets/，平台说明注入 references/平台执行说明.md。

用法：
  python3 build_skill.py --version 1.0.2
  python3 build_skill.py --platform doubao --no-assets   # 开放平台上传版，不含模板
"""
from __future__ import annotations

import argparse
import os
import stat
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
PKG_NAME = 'industrial-robot-solution'
FIXED_TIME = (2020, 1, 1, 0, 0, 0)
PLATFORMS = ('chatgpt', 'doubao')

# 开放平台上传限制，仅用于提示
GATE_TOTAL = 16 * 1024 * 1024
GATE_SINGLE = 2 * 1024 * 1024


def git(*args: str) -> str:
    try:
        return subprocess.run(['git', *args], cwd=ROOT, capture_output=True,
                              text=True, check=True).stdout.strip()
    except (subprocess.CalledProcessError, FileNotFoundError):
        return ''


def collect(platform: str, with_assets: bool) -> list[tuple[Path, str]]:
    """返回 (源文件, 包内相对路径) 列表。"""
    files: list[tuple[Path, str]] = []

    files.append((ROOT / 'skill' / 'SKILL.md', 'SKILL.md'))
    for path in sorted((ROOT / 'skill' / 'references').glob('*.md')):
        files.append((path, f'references/{path.name}'))
    for path in sorted((ROOT / 'skill' / 'scripts').glob('*.py')):
        files.append((path, f'scripts/{path.name}'))

    platform_dir = ROOT / 'platforms' / platform
    doc = platform_dir / '平台执行说明.md'
    if doc.exists():
        files.append((doc, 'references/平台执行说明.md'))
    for path in sorted((platform_dir / 'scripts').glob('*.py')) if (platform_dir / 'scripts').exists() else []:
        files.append((path, f'scripts/{path.name}'))

    if with_assets:
        for path in sorted((ROOT / 'templates').iterdir()):
            if path.is_file() and not path.name.startswith('.'):
                files.append((path, f'assets/{path.name}'))

    missing = [src for src, _ in files if not src.exists()]
    if missing:
        raise SystemExit('缺少源文件：\n' + '\n'.join(str(m) for m in missing))
    return files


def write_zip(target: Path, files: list[tuple[Path, str]], version: str,
              platform: str, commit: str) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    version_note = f'version: {version}\ncommit: {commit or "unknown"}\nplatform: {platform}\n'
    entries = [(src, arc) for src, arc in files]
    with zipfile.ZipFile(target, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        info = zipfile.ZipInfo(f'{PKG_NAME}/VERSION', date_time=FIXED_TIME)
        info.external_attr = (stat.S_IFREG | 0o644) << 16
        zf.writestr(info, version_note)
        for src, arc in sorted(entries, key=lambda item: item[1]):
            info = zipfile.ZipInfo(f'{PKG_NAME}/{arc}', date_time=FIXED_TIME)
            info.external_attr = (stat.S_IFREG | 0o644) << 16
            info.compress_type = zipfile.ZIP_DEFLATED
            zf.writestr(info, src.read_bytes())


def report(target: Path) -> tuple[int, list[tuple[str, int]]]:
    total = 0
    big: list[tuple[str, int]] = []
    with zipfile.ZipFile(target) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            total += info.file_size
            if info.file_size > GATE_SINGLE:
                big.append((info.filename, info.file_size))
    return total, big


def build(platform: str, version: str, out_dir: Path, with_assets: bool) -> Path:
    commit = git('rev-parse', '--short', 'HEAD')
    dirty = git('status', '--porcelain')
    files = collect(platform, with_assets)
    suffix = '' if with_assets else '_no-assets'
    target = out_dir / f'{PKG_NAME}_v{version}_{platform}{suffix}.zip'
    write_zip(target, files, version, platform, commit)
    total, big = report(target)
    print(f'[{platform}] {target.name}  {target.stat().st_size / 1048576:.1f} MB  '
          f'{len(files) + 1} 个文件  解压总量 {total / 1048576:.1f} MB')
    for name, size in big:
        print(f'  ! 单文件超过 2 MB：{name}（{size / 1048576:.1f} MB）'
              '，不能直接上传豆包开放平台（限整包 16 MB、单文件 2 MB）')
    if dirty:
        print('  ! 工作区有未提交改动，本包内容与提交不完全一致')
    return target


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--version', default='1.0.2', help='三段式版本号，写进包内 VERSION')
    parser.add_argument('--platform', default='all', choices=('all',) + PLATFORMS)
    parser.add_argument('--out', default='dist', help='输出目录')
    parser.add_argument('--no-assets', action='store_true',
                        help='不打包模板（用于豆包开放平台上传，模板另行分发）')
    args = parser.parse_args()

    out_dir = (ROOT / args.out).resolve()
    targets = PLATFORMS if args.platform == 'all' else (args.platform,)
    built = [build(p, args.version, out_dir, not args.no_assets) for p in targets]
    print('完成：' + '、'.join(t.name for t in built))


if __name__ == '__main__':
    sys.exit(main())
