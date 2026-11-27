"""生成客户补充信息表 xlsx（可编辑、可回传）。

输入 JSON（列表或 {"issues": [...]}）或 CSV，字段：
  id/category/station/question/impact/unit/basis
中文表头：问题编号、类别、[工位]、请确认的问题、影响的方案判断、客户填写、单位或口径、附件或依据。
多工位项目自动加「工位」列；「客户填写」「附件或依据」留空给客户回填并标浅蓝底色。

用法：
  python3 scripts/make_supplement_xlsx.py --input issues.json --output "客户名_项目名_W01_客户补充信息表.xlsx"
"""
from argparse import ArgumentParser
from pathlib import Path
import csv
import json

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


HEADER_FILL = PatternFill('solid', fgColor='1F3864')
INPUT_FILL = PatternFill('solid', fgColor='DCE6F1')
THIN = Side(style='thin', color='BFBFBF')
WHITE = Side(style='thin', color='FFFFFF')
ALIASES = {
    'id': ('id', '问题编号', '编号'),
    'category': ('category', '类别'),
    'station': ('station', '工位', '工位编号'),
    'question': ('question', '请确认的问题', '问题'),
    'impact': ('impact', '影响的方案判断', '影响'),
    'unit': ('unit', '单位或口径', '单位'),
    'basis': ('basis', '附件或依据', '依据'),
}
WIDTHS = {'问题编号': 10, '类别': 14, '工位': 9, '请确认的问题': 52,
          '影响的方案判断': 40, '客户填写': 30, '单位或口径': 16, '附件或依据': 26}


def normalize(item, index):
    row = {}
    for key, names in ALIASES.items():
        value = next((item[name] for name in names if item.get(name) not in (None, '')), '')
        row[key] = str(value).strip()
    row['id'] = row['id'] or f'Q{index:02d}'
    return row


def load(path):
    text = Path(path).read_text(encoding='utf-8')
    if path.lower().endswith('.json'):
        data = json.loads(text)
        return data.get('issues', data) if isinstance(data, dict) else data
    return list(csv.DictReader(text.splitlines()))


def main():
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--input', required=True)
    parser.add_argument('--output', required=True)
    parser.add_argument('--sheet-name', default='客户补充信息')
    args = parser.parse_args()

    issues = [normalize(item, i) for i, item in enumerate(load(args.input), 1)]
    if not issues:
        raise SystemExit('输入没有问题项，未生成表格。')
    with_station = any(item['station'] for item in issues)

    columns = ['问题编号', '类别'] + (['工位'] if with_station else []) + [
        '请确认的问题', '影响的方案判断', '客户填写', '单位或口径', '附件或依据']
    keys = ['id', 'category'] + (['station'] if with_station else []) + [
        'question', 'impact', None, 'unit', 'basis']
    centered = {'问题编号', '类别', '工位'}
    input_col = columns.index('客户填写') + 1

    wb = Workbook()
    ws = wb.active
    ws.title = args.sheet_name
    ws.append(columns)
    for item in issues:
        ws.append(['' if key is None else item[key] for key in keys])

    for index, name in enumerate(columns, 1):
        cell = ws.cell(row=1, column=index)
        cell.font = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = Border(left=WHITE, right=WHITE, top=WHITE, bottom=WHITE)
        ws.column_dimensions[get_column_letter(index)].width = WIDTHS.get(name, 18)
    ws.row_dimensions[1].height = 28

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(columns)):
        lines = 1
        for cell in row:
            name = columns[cell.column - 1]
            cell.font = Font(name='微软雅黑', size=10)
            cell.alignment = Alignment(
                horizontal='center' if name in centered else 'left',
                vertical='center', wrap_text=True)
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
            if cell.column == input_col:
                cell.fill = INPUT_FILL
            if isinstance(cell.value, str) and cell.value:
                # 列宽以默认字宽计：中日韩字符按 2 个宽度单位估算换行
                units = sum(2 if ord(ch) > 127 else 1 for ch in cell.value)
                capacity = max(4, WIDTHS.get(name, 18) - 1)
                lines = max(lines, -(-units // capacity))
        ws.row_dimensions[row[0].row].height = min(lines * 14 + 8, 120)

    ws.freeze_panes = 'A2'
    ws.sheet_view.zoomScale = 100
    ws.print_title_rows = '1:1'
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    out = Path(args.output).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f'已生成 {out}｜{len(issues)} 个问题｜{"含工位列" if with_station else "项目级"}')

    reopened = load_workbook(out)
    sheet = reopened[args.sheet_name]
    print(f'回读校验：工作表 {reopened.sheetnames}，表头 {[c.value for c in sheet[1]]}，数据行 {sheet.max_row - 1}')


if __name__ == '__main__':
    main()
