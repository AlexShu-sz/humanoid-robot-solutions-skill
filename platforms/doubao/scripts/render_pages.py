"""把 PPTX 转成逐页 PNG，供逐页视觉复核。

用 LibreOffice 转 PDF，再用 pdftoppm 出图。本机没有微软雅黑，渲染结果是替代字体
（Noto Sans CJK）的近似效果：轻微宽度差异可忽略，真实溢出、裁切、遮挡必须处理。

用法：
  python3 scripts/render_pages.py --pptx "<评审稿.pptx>" --out-dir "<工作目录>/render"
  python3 scripts/render_pages.py --pptx "<评审稿.pptx>" --out-dir render --pages 7-9 --dpi 130
"""
from argparse import ArgumentParser
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile


def pages_to_args(pages):
    if not pages:
        return []
    start, sep, end = pages.partition('-')
    start = start.strip() or '1'
    end = end.strip()
    if not sep:
        end = start
    return ['-f', start, '-l', end or start]


def main():
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--pptx', required=True)
    parser.add_argument('--out-dir', required=True)
    parser.add_argument('--dpi', default='110', help='长边约 1500 px 时用 110')
    parser.add_argument('--pages', help='只渲染指定页，例如 7-9 或 12')
    parser.add_argument('--prefix', default='p')
    args = parser.parse_args()

    pptx = Path(args.pptx).expanduser().resolve()
    if not pptx.exists():
        sys.exit(f'找不到文件：{pptx}')
    for tool in ('soffice', 'pdftoppm'):
        if not shutil.which(tool):
            sys.exit(f'缺少命令 {tool}，无法渲染。')

    out_dir = Path(args.out_dir).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    for old in out_dir.glob(f'{args.prefix}-*.png'):
        old.unlink()

    with tempfile.TemporaryDirectory() as tmp:
        proc = subprocess.run(
            ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', tmp, str(pptx)],
            capture_output=True, text=True, timeout=600,
        )
        pdf = next(Path(tmp).glob('*.pdf'), None)
        if proc.returncode != 0 or pdf is None:
            sys.exit(f'LibreOffice 转换失败：{proc.stderr.strip() or proc.stdout.strip()}')

        cmd = ['pdftoppm', '-png', '-r', args.dpi, *pages_to_args(args.pages),
               str(pdf), str(out_dir / f'{args.prefix}')]
        proc = subprocess.run(cmd, capture_output=True, text=True)
        if proc.returncode != 0:
            sys.exit(f'pdftoppm 出图失败：{proc.stderr.strip()}')

    from pptx import Presentation
    slide_count = len(Presentation(str(pptx)).slides)
    images = sorted(out_dir.glob(f'{args.prefix}-*.png'))
    print(f'源文件 {slide_count} 页，生成 {len(images)} 张图：{out_dir}'
          + (f'（仅 {args.pages}）' if args.pages else ''))
    if not args.pages and len(images) != slide_count:
        print('页数与图片数量不一致，先检查文件是否完整。')
    print('逐页看图，重点：文字遮挡、超出文本框、图片裁切、表格行高、断词孤字、页底小字注释。')


if __name__ == '__main__':
    main()
