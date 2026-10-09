import argparse
import copy
import importlib.util
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

ROOT = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('format_check', ROOT / 'check.py')
checker = importlib.util.module_from_spec(spec)
spec.loader.exec_module(checker)
NS = checker.NS


def tag(prefix, name):
    return f'{{{NS[prefix]}}}{name}'


class Package:
    def __init__(self, path):
        with ZipFile(path) as archive:
            self.parts = {name: archive.read(name) for name in archive.namelist() if not name.endswith('/')}
        self.xml_parts = {}

    def xml(self, part):
        if part not in self.xml_parts:
            self.xml_parts[part] = ET.fromstring(self.parts[part])
        return self.xml_parts[part]

    def save(self, path):
        with ZipFile(path, 'w', compression=ZIP_DEFLATED) as archive:
            for name in sorted(self.parts.keys() | self.xml_parts.keys()):
                entry = ZipInfo(name, date_time=(2000, 1, 1, 0, 0, 0))
                entry.compress_type = ZIP_DEFLATED
                archive.writestr(entry, ET.tostring(self.xml_parts[name], encoding='utf-8', xml_declaration=True)
                                 if name in self.xml_parts else self.parts[name])

    def cell(self, address, sheet=1):
        root = self.xml(f'xl/worksheets/sheet{sheet}.xml')
        node = root.find(f'.//s:c[@r="{address}"]', NS)
        if node is None:
            number = ''.join(c for c in address if c.isdigit())
            data = root.find('s:sheetData', NS)
            row = data.find(f's:row[@r="{number}"]', NS)
            if row is None:
                row = ET.SubElement(data, tag('s', 'row'), r=number)
            node = ET.SubElement(row, tag('s', 'c'), r=address)
        return node

    def set_cell(self, address, value, formula=None, sheet=1):
        node = self.cell(address, sheet)
        node[:] = []
        node.attrib.pop('t', None)
        if formula is not None:
            ET.SubElement(node, tag('s', 'f')).text = formula
        if isinstance(value, str):
            node.set('t', 'inlineStr')
            ET.SubElement(ET.SubElement(node, tag('s', 'is')), tag('s', 't')).text = value
        else:
            ET.SubElement(node, tag('s', 'v')).text = str(value)

    def style_cell(self, address, sheet=1, number_format=None, header=False):
        node = self.cell(address, sheet)
        root = self.xml('xl/styles.xml')
        xfs = root.find('s:cellXfs', NS)
        xf = copy.deepcopy(xfs[int(node.get('s', '0'))])
        if number_format:
            formats = root.find('s:numFmts', NS)
            if formats is None:
                formats = ET.Element(tag('s', 'numFmts'), count='0')
                root.insert(0, formats)
            found = next((fmt for fmt in formats if fmt.get('formatCode') == number_format), None)
            if found is None:
                fmt_id = str(max([163] + [int(fmt.attrib['numFmtId']) for fmt in formats]) + 1)
                found = ET.SubElement(formats, tag('s', 'numFmt'), numFmtId=fmt_id, formatCode=number_format)
                formats.set('count', str(len(formats)))
            xf.set('numFmtId', found.attrib['numFmtId'])
            xf.set('applyNumberFormat', '1')
        if header:
            fonts = root.find('s:fonts', NS)
            font = copy.deepcopy(fonts[int(xf.get('fontId', '0'))])
            if font.find('s:b', NS) is None:
                ET.SubElement(font, tag('s', 'b'))
            xf.set('fontId', str(len(fonts)))
            fonts.append(font)
            fonts.set('count', str(len(fonts)))
            fills = root.find('s:fills', NS)
            fill = ET.SubElement(fills, tag('s', 'fill'))
            pattern = ET.SubElement(fill, tag('s', 'patternFill'), patternType='solid')
            ET.SubElement(pattern, tag('s', 'fgColor'), rgb='FFD9EAF7')
            xf.set('fillId', str(len(fills) - 1))
            fills.set('count', str(len(fills)))
        node.set('s', str(len(xfs)))
        xfs.append(xf)
        xfs.set('count', str(len(xfs)))

    def insert_growth_column(self):
        for cell in self.xml('xl/worksheets/sheet1.xml').findall('.//s:c', NS):
            if cell.attrib['r'].startswith('D'):
                cell.set('r', 'E' + cell.attrib['r'][1:])
        self.set_cell('D1', 'Growth %')

    def sort_rows(self, order):
        data = self.xml('xl/worksheets/sheet1.xml').find('s:sheetData', NS)
        original = {int(row.attrib['r']): copy.deepcopy(row) for row in data}
        for row in list(data):
            if int(row.attrib['r']) >= 2:
                data.remove(row)
        for destination, origin in enumerate(order, 2):
            row = original[origin]
            row.set('r', str(destination))
            for cell in row:
                letters = ''.join(c for c in cell.attrib['r'] if c.isalpha())
                cell.set('r', letters + str(destination))
            data.append(row)

    def shape(self, slide, name):
        root = self.xml(f'ppt/slides/slide{slide}.xml')
        return next(shape for shape in root.findall('p:cSld/p:spTree/p:sp', NS)
                    if shape.find('p:nvSpPr/p:cNvPr', NS).get('name') == name)

    def shape_text(self, slide, name, texts):
        body = self.shape(slide, name).find('p:txBody', NS)
        paragraphs = body.findall('a:p', NS)
        template = copy.deepcopy(paragraphs[-1])
        for paragraph in paragraphs:
            body.remove(paragraph)
        for i, text in enumerate(texts):
            paragraph = copy.deepcopy(paragraphs[i] if i < len(paragraphs) else template)
            paragraph.find('a:r/a:t', NS).text = text
            body.append(paragraph)

    def notes(self, slide, texts):
        root = self.xml(f'ppt/notesSlides/notesSlide{slide}.xml')
        body = root.find('p:cSld/p:spTree/p:sp/p:txBody', NS)
        for paragraph in body.findall('a:p', NS):
            body.remove(paragraph)
        for text in texts:
            paragraph = ET.SubElement(body, tag('a', 'p'))
            ET.SubElement(ET.SubElement(paragraph, tag('a', 'r')), tag('a', 't')).text = text


def correct_xlsx(package, number):
    if number == 1:
        package.set_cell('B3', 180)
        package.set_cell('B5', 680, 'SUM(B2:B4)')
    elif number == 2:
        package.set_cell('C5', 630, 'SUM(C2:C4)')
    elif number == 3:
        for sheet, column in [(1, 'A'), (2, 'B')]:
            package.set_cell(column + '2', 'Forest Mix', sheet=sheet)
            package.set_cell(column + '4', 'Forest Mix family pack', sheet=sheet)
    elif number == 4:
        for address in ['A1', 'B1', 'C1']:
            package.style_cell(address, header=True)
    elif number == 5:
        package.sort_rows([3, 5, 2, 4])
    elif number in (6, 9):
        planned, actual = (450, 180) if number == 6 else (350, 470)
        package.set_cell('A5', 'Total')
        package.set_cell('B5', planned, 'SUM(B2:B4)')
        package.set_cell('C5', actual, 'SUM(C2:C4)')
        for address in ['B5', 'C5']:
            package.style_cell(address, number_format='$#,##0.00')
        if number == 9:
            for address, value, formula in [('B2', 350, "'Expenses'!B5"), ('B3', 470, "'Expenses'!C5"), ('B4', 120, 'B3-B2')]:
                package.set_cell(address, value, formula, sheet=2)
                package.style_cell(address, sheet=2, number_format='$#,##0.00')
    elif number in (7, 10):
        package.insert_growth_column()
        if number == 10:
            package.sort_rows([3, 2, 4])
        values = [.25, .5, -.2] if number == 7 else [.5, .25, -.2]
        for row, value in enumerate(values, 2):
            package.set_cell(f'D{row}', value, f'(C{row}-B{row})/B{row}')
            package.style_cell(f'D{row}', number_format='0.0%')
        if number == 10:
            package.set_cell('A5', 'Total')
            package.set_cell('B5', 400, 'SUM(B2:B4)')
            package.set_cell('C5', 505, 'SUM(C2:C4)')
            package.set_cell('D5', .2625, '(C5-B5)/B5')
            for address in ['B5', 'C5']:
                package.style_cell(address, number_format='$#,##0.00')
            package.style_cell('D5', number_format='0.0%')
            for address in ['A1', 'B1', 'C1', 'D1', 'E1']:
                package.style_cell(address, header=True)
    elif number == 8:
        sheets = package.xml('xl/workbook.xml').find('s:sheets', NS)
        entries = list(sheets)
        entries[0].set('name', 'Outlook')
        sheets[:] = [entries[2], entries[0], entries[1]]
        package.set_cell('B2', 450, "'Outlook'!B2", sheet=2)


def correct_pptx(package, number):
    if number == 1:
        package.shape_text(1, 'Title', ['Quarterly results'])
    elif number == 2:
        package.shape_text(2, 'Bullets', ['Revenue: $1.5 million', 'Margin: 18%', 'Retention: 91%'])
    elif number == 3:
        package.shape_text(2, 'Bullets', ['Finish the pilot', 'Collect feedback', 'Launch in October'])
    elif number == 4:
        for slide in [1, 2, 3]:
            for text in package.xml(f'ppt/slides/slide{slide}.xml').findall('.//a:t', NS):
                text.text = text.text.replace('Atlas', 'Beacon')
    elif number == 5:
        package.shape_text(2, 'Title', ['Results'])
        package.shape_text(3, 'Title', ['Next steps'])
    elif number in (6, 7):
        listing = package.xml('ppt/presentation.xml').find('p:sldIdLst', NS)
        original = list(listing)
        listing[:] = [original[i - 1] for i in ([1, 2, 4] if number == 6 else [1, 3, 4, 2])]
    elif number == 8:
        package.notes(2, ['Pause for questions after the chart.', 'Ask which metric needs a follow-up.'])
    elif number == 9:
        shape = copy.deepcopy(package.shape(1, 'Title'))
        identity = shape.find('p:nvSpPr/p:cNvPr', NS)
        identity.set('id', '4')
        identity.set('name', 'Draft label')
        nv = shape.find('p:nvSpPr/p:nvPr', NS)
        nv[:] = []
        transform = shape.find('p:spPr/a:xfrm', NS)
        transform.find('a:off', NS).attrib.update(x='457200', y='4572000')
        transform.find('a:ext', NS).attrib.update(cx='2743200', cy='457200')
        shape.find('p:txBody/a:p/a:r/a:t', NS).text = 'Internal draft'
        style = shape.find('p:txBody/a:p/a:r/a:rPr', NS)
        style.attrib = {'sz': '1800'}
        style.find('a:solidFill/a:srgbClr', NS).set('val', '666666')
        package.xml('ppt/slides/slide1.xml').find('p:cSld/p:spTree', NS).append(shape)
    elif number == 10:
        for folder in ['slides', 'notesSlides']:
            singular = 'slide' if folder == 'slides' else 'notesSlide'
            for path in [f'ppt/{folder}/{singular}2.xml', f'ppt/{folder}/_rels/{singular}2.xml.rels']:
                new_path = path.replace(f'{singular}2', f'{singular}5')
                package.xml_parts[new_path] = copy.deepcopy(package.xml(path))
                if path.endswith('.rels'):
                    for rel in package.xml_parts[new_path]:
                        rel.set('Target', rel.attrib['Target'].replace('slide2.xml', 'slide5.xml').replace('notesSlide2.xml', 'notesSlide5.xml'))
        package.shape_text(5, 'Title', ['Pilot plan'])
        package.shape_text(5, 'Bullets', ['Start with ten users', 'Collect feedback daily', 'Report results in two weeks'])
        package.notes(5, ['Keep the pilot to two weeks.'])
        listing = package.xml('ppt/presentation.xml').find('p:sldIdLst', NS)
        entry = ET.Element(tag('p', 'sldId'), {'id': '999', tag('r', 'id'): 'rIdNew'})
        listing.insert(2, entry)
        links = package.xml('ppt/_rels/presentation.xml.rels')
        ET.SubElement(links, tag('rel', 'Relationship'), Id='rIdNew', Type=NS['r'] + '/slide', Target='slides/slide5.xml')
        types = package.xml('[Content_Types].xml')
        for part, kind in [('ppt/slides/slide5.xml', 'slide'), ('ppt/notesSlides/notesSlide5.xml', 'notesSlide')]:
            ET.SubElement(types, '{http://schemas.openxmlformats.org/package/2006/content-types}Override', PartName='/' + part,
                          ContentType=f'application/vnd.openxmlformats-officedocument.presentationml.{kind}+xml')


def rejects(task, path):
    try:
        checker.check(task, path)
    except Exception:
        return
    raise AssertionError(f'{task["id"]}: invalid output was accepted')


def damage_requested_change(package, format, number):
    if format == 'xlsx':
        if number == 1:
            package.cell('B5').find('s:v', NS).text = '620'
        elif number == 2:
            package.set_cell('C5', 630)
        elif number == 3:
            package.set_cell('B4', 'Trail Mix family pack', sheet=2)
        elif number == 4:
            package.cell('C1').attrib.pop('s')
        elif number == 5:
            package.set_cell('A2', 'Notebook')
        elif number == 6:
            package.cell('B5').attrib.pop('s')
        elif number == 7:
            package.set_cell('D2', .25)
        elif number == 8:
            tabs = package.xml('xl/workbook.xml').find('s:sheets', NS)
            tabs[:] = [tabs[1], tabs[0], tabs[2]]
        elif number == 9:
            package.cell('B3', sheet=2).find('s:f', NS).text = "'Expenses'!B5"
        elif number == 10:
            package.cell('D5').find('s:v', NS).text = '0.25'
    elif number == 1:
        package.shape(1, 'Title').find('p:txBody/a:p/a:r/a:rPr', NS).set('b', '0')
    elif number == 2:
        package.shape_text(2, 'Bullets', ['Revenue: $1.5 million', 'Margin: 19%', 'Retention: 91%'])
    elif number == 3:
        package.shape_text(2, 'Bullets', ['Finish the pilot\nCollect feedback\nLaunch in October'])
    elif number == 4:
        package.shape_text(3, 'Bullets', ['Launch Beacon in October', 'Keep the Atlas team informed'])
    elif number == 5:
        package.shape_text(3, 'Title', ['Draft next steps'])
    elif number == 6:
        listing = package.xml('ppt/presentation.xml').find('p:sldIdLst', NS)
        listing.append(ET.Element(tag('p', 'sldId'), {'id': '258', tag('r', 'id'): 'rId3'}))
    elif number == 7:
        package.notes(2, ['Presenter cue: Results'])
    elif number == 8:
        package.notes(2, ['Pause for questions after the chart.\nAsk which metric needs a follow-up.'])
    elif number == 9:
        shape = package.xml('ppt/slides/slide1.xml').find('p:cSld/p:spTree', NS)[-1]
        shape.find('p:spPr/a:xfrm/a:off', NS).set('y', '4572001')
    elif number == 10:
        for props in package.shape(5, 'Bullets').findall('p:txBody/a:p/a:pPr', NS):
            props[:] = [ET.Element(tag('a', 'buNone'))]


def validate(directory):
    for task in checker.tasks():
        fixture = ROOT / task['fixture']
        package = Package(fixture)
        number = int(task['id'].split('-')[1])
        (correct_xlsx if task['format'] == 'xlsx' else correct_pptx)(package, number)
        output = directory / f'{task["id"]}.{task["format"]}'
        package.save(output)
        checker.check(task, output)
        rejects(task, fixture)
        if task['format'] == 'xlsx':
            package.set_cell('A1', 'Unexpected edit')
        else:
            package.shape_text(1, 'Bullets', ['Unexpected edit', 'Discuss risks'])
        corrupted = directory / f'{task["id"]}-corrupt.{task["format"]}'
        package.save(corrupted)
        rejects(task, corrupted)
        package = Package(output)
        damage_requested_change(package, task['format'], number)
        damaged = directory / f'{task["id"]}-damaged.{task["format"]}'
        package.save(damaged)
        rejects(task, damaged)
        print(f'{task["id"]}: correct PASS; untouched FAIL; unrelated edit FAIL; damaged change FAIL')
    print('20/20 correct outputs accepted; 20/20 untouched fixtures rejected; 40/40 corrupted outputs rejected')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--output-dir', type=Path)
    args = parser.parse_args()
    if args.output_dir:
        args.output_dir.mkdir(parents=True, exist_ok=True)
        validate(args.output_dir)
    else:
        with tempfile.TemporaryDirectory(prefix='office-checkers-') as temporary:
            validate(Path(temporary))


if __name__ == '__main__':
    main()
