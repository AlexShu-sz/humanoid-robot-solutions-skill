"""校验 PPMT 1.0 模板的固定页位与复用格式。

模板默认取技能自带 assets/，未打包的仓库内运行时回落 templates/；
页位缺失时给出可读错误而不是抛出异常，校验项数量与结论一并打印。
"""
from argparse import ArgumentParser
from pathlib import Path
import re
import sys

from pptx import Presentation
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL
from pptx.enum.shapes import MSO_SHAPE_TYPE


SKILL_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_NAME = '人形机器人方案_PPT模板_V1.0.0.pptx'
# 打包后取 assets/，仓库内直接运行回落 templates/
_CANDIDATES = [SKILL_ROOT / 'assets' / TEMPLATE_NAME, SKILL_ROOT.parent / 'templates' / TEMPLATE_NAME]
DEFAULT_TEMPLATE = next((p for p in _CANDIDATES if p.exists()), _CANDIDATES[0])

PRODUCT_TITLE = '机器人本体选型与规格'
ROLES = ('人工保留', '机器人执行', '系统或设备')


def text_shapes(slide):
    return [shape for shape in slide.shapes if shape.has_text_frame and shape.text.strip()]


def named_slide(deck, title):
    for slide in deck.slides:
        if title in {shape.text for shape in text_shapes(slide)}:
            return slide
    return None


def rectangle(shape):
    return tuple(round(getattr(shape, field) / 12700, 1) for field in ('left', 'top', 'width', 'height'))


def run_size(shape):
    run = next(run for paragraph in shape.text_frame.paragraphs for run in paragraph.runs)
    return run.font.size


def signature(shape):
    return shape.text, rectangle(shape), run_size(shape)


def fill_color(shape):
    if shape.fill.type == MSO_FILL.SOLID and shape.fill.fore_color.type == MSO_COLOR_TYPE.RGB:
        return str(shape.fill.fore_color.rgb)
    return None


def bom_structure(deck):
    tables = [shape.table for slide in deck.slides
              if '本项目配置与报价参考' in {shape.text for shape in text_shapes(slide)}
              for shape in slide.shapes if shape.has_table]
    return [
        ([cell.text for cell in table.rows[0].cells], [
            (row.cells[0].text, row.cells[1].text, row.cells[5 if table_index == 0 else 3].text)
            for row in list(table.rows)[1:]
        ])
        for table_index, table in enumerate(tables)
    ]


def check(template, output, cover, product_subtitle, template_timeline, template_scenes, standard_bom):
    source, deck = Presentation(str(template)), Presentation(str(output))
    errors, checks = [], []

    cover_texts = {shape.text for shape in text_shapes(deck.slides[0])}
    checks.append('封面')
    if cover not in cover_texts:
        errors.append(f'封面未逐字采用「{cover}」。当前封面文本：{" / ".join(sorted(cover_texts))[:120]}')

    manual, robot = named_slide(deck, '工人动作拆解'), named_slide(deck, '机器人等效作业方案')
    robot_source = named_slide(source, '机器人等效作业方案')
    if manual is None or robot is None or robot_source is None:
        errors.append('缺少「工人动作拆解」或「机器人等效作业方案」页。')
    else:
        checks.append('动作编号')
        manual_actions = [shape.text for shape in text_shapes(manual) if re.match(r'^\d{2}\s+', shape.text)]
        robot_actions = [shape.text for shape in text_shapes(robot) if re.match(r'^\d{2}\s+', shape.text)]
        if not manual_actions or manual_actions != robot_actions:
            errors.append('工人动作和机器人等效动作的名称、编号或顺序不一致。')

        checks.append('角色标签')
        styles = {shape.text: (fill_color(shape), rectangle(shape)[2:], run_size(shape))
                  for shape in text_shapes(robot_source) if shape.text in ROLES}
        role_fills = {style[0] for style in styles.values()}
        role_labels = [shape for shape in text_shapes(robot)
                       if shape.text in ROLES or fill_color(shape) in role_fills]
        if not role_labels:
            errors.append('机器人等效作业页缺少模板角色标签。')
        for shape in role_labels:
            if (fill_color(shape), rectangle(shape)[2:], run_size(shape)) != styles.get(shape.text):
                errors.append(f'角色标签「{shape.text}」的文字或格式与模板不一致。')

    if template_timeline:
        checks.append('阶段时间轴')
        title = '小步试点，分阶段闭环交付'
        source_timeline, timeline = named_slide(source, title), named_slide(deck, title)
        if source_timeline is None or timeline is None:
            errors.append('缺少「小步试点，分阶段闭环交付」时间轴页。')
        else:
            periods = [signature(shape) for shape in text_shapes(source_timeline)
                       if re.fullmatch(r'\d+–\d+ 周', shape.text)]
            timeline_text = text_shapes(timeline)
            if len(periods) != 5 or not set(periods) <= {signature(shape) for shape in timeline_text}:
                errors.append('缺排期时须沿用模板五个阶段周次及其样式。')
            if not any('计划参考' in shape.text for shape in timeline_text):
                errors.append('模板周次须标为计划参考。')

    checks.append('本体组页位')
    groups = [[slide for slide in item.slides if PRODUCT_TITLE in {shape.text for shape in text_shapes(slide)}]
              for item in (source, deck)]
    if any(len(group) < 3 for group in groups):
        errors.append('缺少本体规格、应用示例或适配评估页。')
        return errors, checks

    checks.append('规格页副标题')
    source_specs, source_scenes = groups[0][:2]
    specs, scenes = groups[1][:2]
    old_subtitle = next((shape for shape in text_shapes(source_specs) if 'Cruzr S2' in shape.text), None)
    if old_subtitle is None:
        errors.append('模板规格页未找到型号副标题，确认模板是否为 PPMT 1.0。')
    else:
        expected = (product_subtitle or old_subtitle.text, rectangle(old_subtitle), run_size(old_subtitle))
        if expected not in {signature(shape) for shape in text_shapes(specs)}:
            errors.append('本体规格页副标题应保留正式型号和产品类型的模板格式。')

    checks.append('场景页标题与图注')
    if not {signature(shape) for shape in text_shapes(source_scenes)} <= {signature(shape) for shape in text_shapes(scenes)}:
        errors.append('场景页缺少模板主副标题或对应图注，或其位置、字号有变化。')

    if template_scenes:
        checks.append('场景图原图')
        source_images = [shape for shape in source_scenes.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
        scene_images = [shape for shape in scenes.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
        if [(rectangle(shape), shape.image.sha1) for shape in source_images] != [
            (rectangle(shape), shape.image.sha1) for shape in scene_images
        ]:
            errors.append('模板工业应用场景图发生替换、缺失或移位。')

    if standard_bom:
        checks.append('BOM 结构与编号')
        reference_bom, output_bom = bom_structure(source), bom_structure(deck)
        if len(reference_bom) != 2 or output_bom != reference_bom:
            errors.append('配置与报价表头、标准项目名称、顺序或单位与模板不一致。')
        else:
            software_rows = output_bom[1][1]
            software_start = next((i for i, row in enumerate(software_rows) if row[0] == '三、软件'), None)
            if software_start is None or [row[1] for row in software_rows[software_start + 1:software_start + 3]] != ['软件和场景开发', '']:
                errors.append('软件组仅保留「软件和场景开发」。')
            numbers = [number for _, rows in output_bom for number, _, _ in rows if number.isdigit()]
            if numbers != [f'{i:02}' for i in range(1, len(numbers) + 1)]:
                errors.append('配置及报价的标准项目编号须跨页连续。')
    return errors, checks


def main():
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--template', default=str(DEFAULT_TEMPLATE), help='模板 PPTX，默认取技能自带资产')
    parser.add_argument('--output', required=True, help='待校验的评审稿 PPTX')
    parser.add_argument('--cover', required=True, help='封面应逐字出现的客户名与项目行')
    parser.add_argument('--product-subtitle', help='已确认更换本体时填完整正式型号和产品类型')
    parser.add_argument('--template-timeline', action='store_true')
    parser.add_argument('--template-scenes', action='store_true')
    parser.add_argument('--standard-bom', action='store_true')
    args = parser.parse_args()

    for path in (args.template, args.output):
        if not Path(path).exists():
            parser.exit(2, f'文件不存在：{path}\n')
    try:
        errors, checks = check(args.template, args.output, args.cover, args.product_subtitle,
                               args.template_timeline, args.template_scenes, args.standard_bom)
    except Exception as exc:  # 模板与页位不匹配时给出可读提示
        parser.exit(2, f'校验无法完成：{type(exc).__name__}: {exc}\n'
                       '确认 --template 与 --output 是同一套模板生成的 PPTX，且文件可正常打开。\n')

    print('已检查：' + '、'.join(checks))
    if errors:
        parser.exit(1, '\n'.join(errors) + '\n')
    print('PPMT 1.0 固定页位校验通过。')


if __name__ == '__main__':
    main()
