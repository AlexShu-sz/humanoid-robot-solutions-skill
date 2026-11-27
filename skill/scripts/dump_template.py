"""导出模板或评审稿的页位与形状清单，供定位填充位置。

先看页级清单（默认），再用 --page 看单页形状：名称、类型、位置尺寸（cm）、文字预览、
表格行列、图片尺寸。填模板前先跑一次，不要凭页号猜形状。
"""
from argparse import ArgumentParser
from pathlib import Path
import json

from pptx import Presentation
from pptx.util import Emu


def cm(value):
    return round(Emu(value).cm, 2)


def text_of(shape):
    return ' | '.join(part.strip() for part in shape.text_frame.text.split('\n') if part.strip())


def page_summary(index, slide):
    texts = [text_of(shape) for shape in slide.shapes if shape.has_text_frame and shape.text.strip()]
    pictures = sum(1 for shape in slide.shapes if shape.shape_type == 13)
    tables = sum(1 for shape in slide.shapes if shape.has_table)
    return {
        'page': index,
        'texts': texts[:4],
        'shapes': len(slide.shapes),
        'pictures': pictures,
        'tables': tables,
    }


def page_detail(slide):
    rows = []
    for shape in slide.shapes:
        item = {
            'name': shape.name,
            'type': str(shape.shape_type).split(' ')[0] if shape.shape_type is not None else '?',
            'id': shape.shape_id,
            'box': [cm(shape.left), cm(shape.top), cm(shape.width), cm(shape.height)],
        }
        if shape.has_text_frame and shape.text.strip():
            item['text'] = text_of(shape)[:80]
        if shape.has_table:
            item['table'] = [len(shape.table.rows), len(shape.table.columns)]
        if shape.shape_type == 13:
            item['picture'] = list(shape.image.size)
        rows.append(item)
    return rows


def main():
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('pptx')
    parser.add_argument('--page', type=int, help='只看某一页（从 1 起）')
    parser.add_argument('--json', action='store_true')
    args = parser.parse_args()

    path = Path(args.pptx).expanduser().resolve()
    if not path.exists():
        raise SystemExit(f'找不到文件：{path}')
    prs = Presentation(str(path))
    print(f'{path.name}｜{len(prs.slides)} 页｜{round(prs.slide_width / 914400, 2)} x {round(prs.slide_height / 914400, 2)} in')

    if args.page:
        slide = prs.slides[args.page - 1]
        detail = page_detail(slide)
        if args.json:
            print(json.dumps(detail, ensure_ascii=False, indent=2))
        else:
            for item in detail:
                extra = f' 表格{item["table"]}行x列' if 'table' in item else ''
                extra += f' 图{item["picture"][0]}x{item["picture"][1]}' if 'picture' in item else ''
                box = ','.join(str(v) for v in item['box'])
                print(f'  {item["id"]:>3} {item["type"]:<10} {box:<22} {item["name"]}｜{item.get("text", "")}{extra}')
        return

    summaries = [page_summary(i, slide) for i, slide in enumerate(prs.slides, 1)]
    if args.json:
        print(json.dumps(summaries, ensure_ascii=False, indent=2))
        return
    for item in summaries:
        head = ' / '.join(item['texts'])
        print(f'P{item["page"]:02d} 形状{item["shapes"]:>2} 图{item["pictures"]} 表{item["tables"]}｜{head[:100]}')


if __name__ == '__main__':
    main()
