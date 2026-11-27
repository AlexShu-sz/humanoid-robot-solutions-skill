"""Check the fixed PPMT 1.0 page formats against the supplied native template."""
from argparse import ArgumentParser
import re

from pptx import Presentation
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL
from pptx.enum.shapes import MSO_SHAPE_TYPE


PRODUCT_TITLE = '机器人本体选型与规格'
ROLES = ('人工保留', '机器人执行', '系统或设备')


def named_slide(deck, title):
    return next(slide for slide in deck.slides if title in {shape.text for shape in text_shapes(slide)})


def text_shapes(slide):
    return [shape for shape in slide.shapes if shape.has_text_frame and shape.text.strip()]


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
        ([cell.text for cell in table.rows[0].cells],[
            (row.cells[0].text,row.cells[1].text,row.cells[5 if table_index == 0 else 3].text)
            for row in list(table.rows)[1:]
        ])
        for table_index, table in enumerate(tables)
    ]


def check(template, output, cover, product_subtitle, template_timeline, template_scenes, standard_bom):
    source, deck = Presentation(template), Presentation(output)
    errors = []
    if cover not in {shape.text for shape in text_shapes(deck.slides[0])}:
        errors.append('封面未逐字采用客户名称与指定展示名。')

    manual = named_slide(deck, '工人动作拆解')
    robot_source = named_slide(source, '机器人等效作业方案')
    robot = named_slide(deck, '机器人等效作业方案')
    manual_actions = [shape.text for shape in text_shapes(manual) if re.match(r'^\d{2}\s+',shape.text)]
    robot_actions = [shape.text for shape in text_shapes(robot) if re.match(r'^\d{2}\s+',shape.text)]
    if not manual_actions or manual_actions != robot_actions:
        errors.append('工人动作和机器人等效动作的名称、编号或顺序不一致。')
    styles = {shape.text: (fill_color(shape),rectangle(shape)[2:],run_size(shape))
              for shape in text_shapes(robot_source) if shape.text in ROLES}
    role_fills = {style[0] for style in styles.values()}
    role_labels = [shape for shape in text_shapes(robot)
                   if shape.text in ROLES or fill_color(shape) in role_fills]
    if not role_labels:
        errors.append('机器人等效作业页缺少模板角色标签。')
    for shape in role_labels:
        if (fill_color(shape),rectangle(shape)[2:],run_size(shape)) != styles.get(shape.text):
            errors.append(f'角色标签「{shape.text}」的文字或格式与模板不一致。')

    if template_timeline:
        title = '小步试点，分阶段闭环交付'
        source_timeline, timeline = [named_slide(item,title) for item in (source,deck)]
        periods = [signature(shape) for shape in text_shapes(source_timeline)
                   if re.fullmatch(r'\d+–\d+ 周',shape.text)]
        timeline_text = text_shapes(timeline)
        if len(periods) != 5 or not set(periods) <= {signature(shape) for shape in timeline_text}:
            errors.append('缺排期时须沿用模板五个阶段周次及其样式。')
        if not any('计划参考' in shape.text for shape in timeline_text):
            errors.append('模板周次须标为计划参考。')

    groups = [[slide for slide in item.slides if PRODUCT_TITLE in {shape.text for shape in text_shapes(slide)}]
              for item in (source,deck)]
    if any(len(group) < 3 for group in groups):
        return errors + ['缺少本体规格、应用示例或适配评估页。']
    source_specs, source_scenes = groups[0][:2]
    specs, scenes = groups[1][:2]
    old_subtitle = next(shape for shape in text_shapes(source_specs) if 'Cruzr S2' in shape.text)
    expected_subtitle = (product_subtitle or old_subtitle.text, rectangle(old_subtitle),run_size(old_subtitle))
    if expected_subtitle not in {signature(shape) for shape in text_shapes(specs)}:
        errors.append('本体规格页副标题应保留正式型号和产品类型的模板格式。')
    if not {signature(shape) for shape in text_shapes(source_scenes)} <= {signature(shape) for shape in text_shapes(scenes)}:
        errors.append('场景页缺少模板主副标题或对应图注，或其位置、字号有变化。')
    source_images, scene_images = [[shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
                                   for slide in (source_scenes,scenes)]
    if template_scenes and [(rectangle(shape),shape.image.sha1) for shape in source_images] != [
        (rectangle(shape),shape.image.sha1) for shape in scene_images
    ]:
        errors.append('模板工业应用场景图发生替换、缺失或移位。')
    if standard_bom:
        reference_bom, output_bom = bom_structure(source), bom_structure(deck)
        if len(reference_bom) != 2 or output_bom != reference_bom:
            errors.append('配置与报价表头、标准项目名称、顺序或单位与模板不一致。')
        else:
            software_rows = output_bom[1][1]
            software_start = next((i for i,row in enumerate(software_rows) if row[0]=='三、软件'),None)
            if software_start is None or [row[1] for row in software_rows[software_start+1:software_start+3]] != ['软件和场景开发','']:
                errors.append('软件组仅保留「软件和场景开发」。')
            numbers = [number for _, rows in output_bom for number, _, _ in rows if number.isdigit()]
            if numbers != [f'{i:02}' for i in range(1,len(numbers)+1)]:
                errors.append('配置及报价的标准项目编号须跨页连续。')
    return errors


def main():
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--template',required=True)
    parser.add_argument('--output',required=True)
    parser.add_argument('--cover',required=True)
    parser.add_argument('--product-subtitle',help='已确认更换本体时填写完整正式型号和产品类型')
    parser.add_argument('--template-timeline',action='store_true')
    parser.add_argument('--template-scenes',action='store_true')
    parser.add_argument('--standard-bom',action='store_true')
    args = parser.parse_args()
    errors = check(args.template,args.output,args.cover,args.product_subtitle,args.template_timeline,args.template_scenes,args.standard_bom)
    if errors:
        parser.exit(1,'\n'.join(errors)+'\n')
    print('PPMT 1.0 固定页位校验通过。')


if __name__ == '__main__':
    main()
