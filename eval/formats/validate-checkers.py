import argparse
import contextlib
import copy
import importlib.util
import tempfile
import xml.etree.ElementTree as ET
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

ROOT = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('format_check', ROOT / 'check.py')
checker = importlib.util.module_from_spec(spec)
spec.loader.exec_module(checker)
NS = {**checker.NS, 'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
      'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006'}
for prefix in ('w', 'r', 'w14', 'mc'):
    ET.register_namespace(prefix, NS[prefix])
ACCEPTED, UNTOUCHED, CORRUPTED = 'correct outputs accepted', 'untouched fixtures rejected', 'corrupted outputs rejected'
COUNTS = {kind: [0, 0] for kind in (ACCEPTED, UNTOUCHED, CORRUPTED)}
FAILURES = []


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
        root = package.xml('xl/workbook.xml')
        sheets = root.find('s:sheets', NS)
        entries = list(sheets)
        entries[0].set('name', 'Outlook')
        sheets[:] = [entries[2], entries[0], entries[1]]
        for view in root.findall('s:bookViews/s:workbookView', NS):
            for key in ('activeTab', 'firstSheet'):
                view.set(key, str(list(sheets).index(entries[int(view.get(key, '0'))])))
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


DOCUMENT = 'word/document.xml'
PILOT = 'The pilot started in September with ten participants across two sites.'
NOTE = 'Draft note: confirm the travel figures with finance before sending.'


def word_body(package):
    return package.xml(DOCUMENT).find('w:body', NS)


def word_text(paragraph):
    return ''.join(node.text or '' for node in paragraph.iter(tag('w', 't')))


def word_paragraph(package, text):
    return next(paragraph for paragraph in package.xml(DOCUMENT).iter(tag('w', 'p')) if word_text(paragraph) == text)


def set_word_text(paragraph, text):
    paragraph.find('w:r/w:t', NS).text = text


def word_rows(package):
    return word_body(package).find('w:tbl', NS).findall('w:tr', NS)


def word_cell(package, row, column):
    return word_rows(package)[row - 1].findall('w:tc', NS)[column - 1].find('w:p', NS)


def word_style(paragraph, style):
    props = paragraph.find('w:pPr', NS)
    if props is None:
        props = ET.Element(tag('w', 'pPr'))
        paragraph.insert(0, props)
    node = props.find('w:pStyle', NS)
    if node is None:
        node = ET.Element(tag('w', 'pStyle'))
        props.insert(0, node)
    node.set(tag('w', 'val'), style)


def remove_word_paragraph(package, text):
    word_body(package).remove(word_paragraph(package, text))


def replace_word_text(package, old, new):
    for node in package.xml(DOCUMENT).iter(tag('w', 't')):
        node.text = node.text.replace(old, new)


def bold(run):
    props = run.find('w:rPr', NS)
    if props is None:
        props = ET.Element(tag('w', 'rPr'))
        run.insert(0, props)
    props.insert(0, ET.Element(tag('w', 'b')))


def split_run(paragraph, pieces):
    run = paragraph.find('w:r', NS)
    runs = []
    for text, edit in pieces:
        piece = copy.deepcopy(run)
        node = piece.find('w:t', NS)
        node.text = text
        if text and (text[0].isspace() or text[-1].isspace()):
            node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        if edit:
            edit(piece)
        runs.append(piece)
    index = list(paragraph).index(run)
    paragraph[index:index + 1] = runs


def insert_word_paragraph(package, after, text):
    paragraph = ET.Element(tag('w', 'p'))
    ET.SubElement(ET.SubElement(paragraph, tag('w', 'r')), tag('w', 't')).text = text
    body = word_body(package)
    body.insert(list(body).index(word_paragraph(package, after)) + 1, paragraph)
    return paragraph


def fresh_word_ids(package, paragraphs):
    used = {int(value, 16) for paragraph in package.xml(DOCUMENT).iter(tag('w', 'p'))
            for name in ('paraId', 'textId') if (value := paragraph.get(tag('w14', name))) is not None}
    value = 1
    for paragraph in paragraphs:
        for name in ('paraId', 'textId'):
            while value in used:
                value += 1
            assert value < 0x80000000
            paragraph.set(tag('w14', name), f'{value:08X}')
            used.add(value)


def correct_docx(package, number):
    if number == 1:
        set_word_text(word_paragraph(package, 'Sumary'), 'Summary')
    elif number == 2:
        replace_word_text(package, 'Atlas', 'Beacon')
    elif number == 3:
        set_word_text(word_cell(package, 3, 4), '130')
        set_word_text(word_cell(package, 5, 4), '660')
    elif number in (4, 5):
        for column, text in zip((3, 4), ('700', '705') if number == 4 else ('500', '530')):
            set_word_text(word_cell(package, 5, column), text)
        table, rows = word_body(package).find('w:tbl', NS), word_rows(package)
        if number == 4:
            row = copy.deepcopy(rows[3])
            fresh_word_ids(package, row.iter(tag('w', 'p')))
            for cell, text in zip(row.findall('w:tc', NS), ['Training', 'Avery', '80', '75']):
                set_word_text(cell.find('w:p', NS), text)
            table.insert(list(table).index(rows[4]), row)
        else:
            table.remove(rows[2])
    elif number == 6:
        word_style(word_paragraph(package, 'Risks and mitigations'), 'Heading2')
    elif number == 7:
        split_run(word_paragraph(package, PILOT), [('The pilot started in September with ', None), ('ten participants', bold),
                                                   (' across two sites.', None)])
    elif number == 8:
        insert_word_paragraph(package, 'Next steps', 'The steering group meets on 20 October.')
    elif number == 9:
        remove_word_paragraph(package, NOTE)
    elif number == 10:
        replace_word_text(package, '[DATE]', '14 October 2026')
        remove_word_paragraph(package, 'DRAFT')
        word_style(word_paragraph(package, 'Pilot status report'), 'Heading1')


CORRECT = {'xlsx': correct_xlsx, 'pptx': correct_pptx, 'docx': correct_docx}


def count(kind):
    COUNTS[kind][0] += 1
    COUNTS[kind][1] += 1


def accepts(task, path):
    checker.check(task, path)
    count(ACCEPTED)


def rejects(task, path, kind=CORRUPTED):
    try:
        checker.check(task, path)
    except Exception:
        count(kind)
        return
    raise AssertionError(f'{task["id"]}: invalid output was accepted')


def expect(name, task, path, accept, source=None, reason=None):
    try:
        with patch.object(checker, 'fixture', return_value=source) if source else contextlib.nullcontext():
            checker.check(task, path)
        outcome = 'accepted'
    except Exception as error:
        outcome = f'rejected ({error})'
    passed = (outcome == 'accepted') == accept and (reason is None or reason in outcome)
    kind = ACCEPTED if accept else CORRUPTED
    COUNTS[kind][0] += passed
    COUNTS[kind][1] += 1
    if not passed:
        FAILURES.append(f'{name}: {outcome}')
    print(f'{name}: {outcome} {"PASS" if passed else "FAIL"}')


def task_named(task_id):
    return next(task for task in checker.tasks() if task['id'] == task_id)


def variant(directory, name, base, *edits):
    package = Package(base)
    for edit in edits:
        edit(package)
    path = directory / f'{name}{Path(base).suffix}'
    package.save(path)
    return path


def select(sheet, value):
    def edit(package):
        for view in package.xml(f'xl/worksheets/sheet{sheet}.xml').findall('s:sheetViews/s:sheetView', NS):
            if value is None:
                view.attrib.pop('tabSelected', None)
            else:
                view.set('tabSelected', value)
    return edit


def move_part(package, old, new):
    for before, after in [(old, new), (checker.relpath(old), checker.relpath(new))]:
        package.xml_parts[after] = package.xml(before)
        package.xml_parts.pop(before)
        package.parts.pop(before)
    for name in package.parts.keys() | package.xml_parts.keys():
        if name.endswith('.rels'):
            for rel in package.xml(name):
                if rel.get('Target', '').endswith('/' + old.rsplit('/', 1)[-1]):
                    rel.set('Target', rel.attrib['Target'].rsplit('/', 1)[0] + '/' + new.rsplit('/', 1)[-1])
    for item in package.xml('[Content_Types].xml'):
        if item.get('PartName') == '/' + old:
            item.set('PartName', '/' + new)


def rename_notes(package):
    move_part(package, 'ppt/notesSlides/notesSlide2.xml', 'ppt/notesSlides/speakerNotes2.xml')


def bold_notes(package):
    for run in package.xml('ppt/notesSlides/notesSlide2.xml').iter(tag('a', 'r')):
        run.insert(0, ET.Element(tag('a', 'rPr'), b='1'))


def notes_owner(notes, slide):
    def edit(package):
        package.xml(f'ppt/notesSlides/_rels/notesSlide{notes}.xml.rels')[0].set('Target', f'../slides/slide{slide}.xml')
    return edit


def notes_content_type(package):
    for item in package.xml('[Content_Types].xml'):
        if item.get('PartName') == '/ppt/notesSlides/notesSlide2.xml':
            item.set('ContentType', 'application/octet-stream')


def notes_geometry(package):
    package.xml('ppt/notesSlides/notesSlide3.xml').find('.//a:off', NS).set('x', '685801')


def without_notes(package):
    part = 'ppt/notesSlides/notesSlide2.xml'
    for name in (part, checker.relpath(part)):
        package.parts.pop(name, None)
        package.xml_parts.pop(name, None)
    links = package.xml('ppt/slides/_rels/slide2.xml.rels')
    for rel in list(links):
        if rel.get('Type', '').endswith('/notesSlide'):
            links.remove(rel)
    types = package.xml('[Content_Types].xml')
    for item in list(types):
        if item.get('PartName') == '/' + part:
            types.remove(item)


def without_notes_links(package):
    name = checker.relpath('ppt/notesSlides/notesSlide2.xml')
    package.parts.pop(name)
    package.xml_parts.pop(name, None)


def renumber_pools(package):
    root = package.xml('xl/styles.xml')
    for pool, attribute, extra in [('fonts', 'fontId', f'<font xmlns="{NS["s"]}"><i/><sz val="9"/><name val="Unused"/></font>'),
                                   ('fills', 'fillId', f'<fill xmlns="{NS["s"]}"><patternFill patternType="lightGray"/></fill>')]:
        entries = root.find(f's:{pool}', NS)
        entries.insert(1, ET.fromstring(extra))
        entries.set('count', str(len(entries)))
        for xf in root.findall('s:cellStyleXfs/s:xf', NS) + root.findall('s:cellXfs/s:xf', NS):
            if int(xf.get(attribute, '0')) >= 1:
                xf.set(attribute, str(int(xf.get(attribute)) + 1))


def engine_header(package):
    root = package.xml('xl/styles.xml')
    fonts, fills, xfs = (root.find(f's:{name}', NS) for name in ('fonts', 'fills', 'cellXfs'))
    font = copy.deepcopy(fonts[0])
    font.insert(0, ET.Element(tag('s', 'b')))
    fonts.append(font)
    fill = ET.SubElement(fills, tag('s', 'fill'))
    ET.SubElement(ET.SubElement(fill, tag('s', 'patternFill'), patternType='solid'), tag('s', 'fgColor'), rgb='FFD9EAF7')
    ET.SubElement(xfs, tag('s', 'xf'), numFmtId='0', fontId=str(len(fonts) - 1), fillId=str(len(fills) - 1), borderId='0',
                  applyFont='1', applyFill='1', applyBorder='1')
    for pool in (fonts, fills, xfs):
        pool.set('count', str(len(pool)))
    for address in ('A1', 'B1', 'C1'):
        package.cell(address).set('s', str(len(xfs) - 1))


def fill_background(pattern_type, **color):
    def edit(package):
        for pattern in package.xml('xl/styles.xml').findall('s:fills/s:fill/s:patternFill', NS):
            if pattern.get('patternType') == pattern_type:
                ET.SubElement(pattern, tag('s', 'bgColor'), **color)
    return edit


def quote_prefix(package):
    package.xml('xl/styles.xml').find('s:cellXfs', NS)[0].set('quotePrefix', '1')


def header_font_not_applied(package):
    xfs = package.xml('xl/styles.xml').find('s:cellXfs', NS)
    xfs[int(package.cell('A1').get('s'))].set('applyFont', '0')


def patterned_cell(package):
    root = package.xml('xl/styles.xml')
    fills, xfs = root.find('s:fills', NS), root.find('s:cellXfs', NS)
    fill = ET.SubElement(fills, tag('s', 'fill'))
    ET.SubElement(ET.SubElement(fill, tag('s', 'patternFill'), patternType='darkGray'), tag('s', 'fgColor'), rgb='FF808080')
    ET.SubElement(xfs, tag('s', 'xf'), numFmtId='0', fontId='0', fillId=str(len(fills) - 1), borderId='0', xfId='0', applyFill='1')
    for pool in (fills, xfs):
        pool.set('count', str(len(pool)))
    package.cell('B1').set('s', str(len(xfs) - 1))


def style_references(root):
    return root.findall('s:cellStyleXfs/s:xf', NS) + root.findall('s:cellXfs/s:xf', NS)


def unused_first(pool, attribute, entry):
    def edit(package):
        root = package.xml('xl/styles.xml')
        entries = root.find(f's:{pool}', NS)
        entries.insert(0, ET.fromstring(entry))
        entries.set('count', str(len(entries)))
        for xf in style_references(root):
            xf.set(attribute, str(int(xf.get(attribute, '0')) + 1))
    return edit


def renumber_formats(package):
    root = package.xml('xl/styles.xml')
    formats = root.find('s:numFmts', NS)
    for item in formats:
        item.set('numFmtId', str(int(item.attrib['numFmtId']) + 10))
    for xf in style_references(root):
        if int(xf.get('numFmtId', '0')) >= 164:
            xf.set('numFmtId', str(int(xf.attrib['numFmtId']) + 10))
    ET.SubElement(formats, tag('s', 'numFmt'), numFmtId='164', formatCode='0.000')
    formats.set('count', str(len(formats)))


def named_heading(package):
    root = package.xml('xl/styles.xml')
    fonts, named, xfs, styles = (root.find(f's:{name}', NS) for name in ('fonts', 'cellStyleXfs', 'cellXfs', 'cellStyles'))
    fonts.append(ET.fromstring(f'<font xmlns="{NS["s"]}"><b/><sz val="13"/><name val="DejaVu Sans"/></font>'))
    font = str(len(fonts) - 1)
    ET.SubElement(named, tag('s', 'xf'), numFmtId='0', fontId=font, fillId='0', borderId='0', applyFont='1')
    ET.SubElement(xfs, tag('s', 'xf'), numFmtId='0', fontId=font, fillId='0', borderId='0', xfId='1')
    ET.SubElement(styles, tag('s', 'cellStyle'), name='Heading 1', xfId='1', builtinId='16')
    for pool in (fonts, named, xfs, styles):
        pool.set('count', str(len(pool)))
    package.cell('A4').set('s', str(len(xfs) - 1))
    sheet = package.xml('xl/worksheets/sheet1.xml')
    columns = ET.Element(tag('s', 'cols'))
    ET.SubElement(columns, tag('s', 'col'), min='5', max='5', width='12', customWidth='1', style='1')
    sheet.insert(list(sheet).index(sheet.find('s:sheetData', NS)), columns)
    sheet.find('s:sheetData/s:row[@r="4"]', NS).attrib.update(s='1', customFormat='1')


def swap_named_formats(package):
    root = package.xml('xl/styles.xml')
    named = root.find('s:cellStyleXfs', NS)
    named[:] = [named[1], named[0]]
    for item in root.findall('s:cellXfs/s:xf', NS) + root.findall('s:cellStyles/s:cellStyle', NS):
        item.set('xfId', str(1 - int(item.get('xfId', '0'))))


def swap_cell_styles(package):
    styles = package.xml('xl/styles.xml').find('s:cellStyles', NS)
    styles[:] = [styles[1], styles[0]]


def move_cell_format(package):
    xfs = package.xml('xl/styles.xml').find('s:cellXfs', NS)
    order = [0, *range(2, len(xfs)), 1]
    moved = {str(old): str(new) for new, old in enumerate(order)}
    xfs[:] = [xfs[old] for old in order]
    sheet = package.xml('xl/worksheets/sheet1.xml')
    for kind, key in (('c', 's'), ('row', 's'), ('col', 'style')):
        for node in sheet.iter(tag('s', kind)):
            if key in node.attrib:
                node.set(key, moved[node.attrib[key]])


def restyle_column(package):
    package.xml('xl/worksheets/sheet1.xml').find('s:cols/s:col', NS).set('style', '0')


def restyle_row(package):
    package.xml('xl/worksheets/sheet1.xml').find('s:sheetData/s:row[@r="4"]', NS).set('s', '0')


def named_font(package):
    root = package.xml('xl/styles.xml')
    fonts = root.find('s:fonts', NS)
    fonts.append(ET.fromstring(f'<font xmlns="{NS["s"]}"><b/><i/><sz val="13"/><name val="DejaVu Sans"/></font>'))
    fonts.set('count', str(len(fonts)))
    root.find('s:cellStyleXfs', NS)[1].set('fontId', str(len(fonts) - 1))


def reorder_children(part, path):
    def edit(package):
        for node in package.xml(part).findall(path, NS):
            node[:] = reversed(node)
    return edit


def named_font_child(name, edit):
    def change(package):
        root = package.xml('xl/styles.xml')
        xf = root.find('s:cellStyleXfs', NS)[1]
        font = root.find('s:fonts', NS)[int(xf.get('fontId', '0'))]
        edit(font, font.find(f's:{name}', NS))
    return change


def empty_cached_value(kind):
    def edit(package):
        cell = package.cell('B5')
        if kind is None:
            cell.attrib.pop('t', None)
        else:
            cell.set('t', kind)
        cell.find('s:v', NS).text = None
    return edit


def rich_comment(package):
    part = 'xl/checker-comments.xml'
    root = ET.Element(tag('s', 'comments'))
    ET.SubElement(ET.SubElement(root, tag('s', 'authors')), tag('s', 'author')).text = 'Reviewer'
    comment = ET.SubElement(ET.SubElement(root, tag('s', 'commentList')), tag('s', 'comment'), ref='A1', authorId='0')
    run = ET.SubElement(ET.SubElement(comment, tag('s', 'text')), tag('s', 'r'))
    properties = ET.SubElement(run, tag('s', 'rPr'))
    ET.SubElement(properties, tag('s', 'sz'), val='11')
    ET.SubElement(properties, tag('s', 'rFont'), val='DejaVu Sans')
    ET.SubElement(run, tag('s', 't')).text = 'Review note'
    package.xml_parts[part] = root
    ET.SubElement(package.xml('[Content_Types].xml'), '{http://schemas.openxmlformats.org/package/2006/content-types}Override',
                  PartName='/' + part, ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml')
    path = 'xl/worksheets/_rels/sheet1.xml.rels'
    links = package.xml(path) if path in package.parts else ET.Element(tag('rel', 'Relationships'))
    ET.SubElement(links, tag('rel', 'Relationship'), Id='rIdCheckerComment', Type=NS['r'] + '/comments', Target='../checker-comments.xml')
    package.xml_parts[path] = links


def comment_properties(edit):
    def change(package):
        edit(package.xml('xl/checker-comments.xml').find('.//s:rPr', NS))
    return change


def cell_style(key, value):
    def edit(package):
        package.xml('xl/styles.xml').find('s:cellStyles', NS)[1].set(key, value)
    return edit


def named_attribute(package):
    package.xml('xl/styles.xml').find('s:cellStyleXfs', NS)[1].set('pivotButton', '1')


def named_child(package):
    xf = package.xml('xl/styles.xml').find('s:cellStyleXfs', NS)[1]
    ET.SubElement(ET.SubElement(xf, tag('s', 'extLst')), tag('s', 'ext'), uri='changed')


def end_properties(part):
    def edit(package):
        paragraph = package.xml(part).find('p:cSld/p:spTree/p:sp/p:txBody/a:p', NS)
        ET.SubElement(paragraph, tag('a', 'endParaRPr'), sz='4800', b='1')
    return edit


def ended_slides(package):
    for slide in (1, 2):
        for paragraph in list(package.xml(f'ppt/slides/slide{slide}.xml').iter(tag('a', 'p'))):
            ET.SubElement(paragraph, tag('a', 'endParaRPr'), sz='1800')
    body = package.shape(2, 'Bullets').find('p:txBody', NS)
    paragraph = copy.deepcopy(body.findall('a:p', NS)[-1])
    run = paragraph.find('a:r', NS)
    run.find('a:t', NS).text = 'Slide '
    field = ET.Element(tag('a', 'fld'), id='{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}', type='slidenum')
    field.append(copy.deepcopy(run.find('a:rPr', NS)))
    ET.SubElement(field, tag('a', 't')).text = '2'
    paragraph.insert(list(paragraph).index(run) + 1, field)
    body.append(paragraph)
    ET.SubElement(package.xml('ppt/notesSlides/notesSlide3.xml').find('.//a:p', NS), tag('a', 'endParaRPr'), sz='1200')


def field_paragraph(package):
    return package.shape(2, 'Bullets').findall('p:txBody/a:p', NS)[-1]


def changed_end(package):
    package.shape(2, 'Bullets').find('p:txBody/a:p/a:endParaRPr', NS).set('sz', '4800')


def dropped_title_end(package):
    paragraph = package.shape(1, 'Title').find('p:txBody/a:p', NS)
    paragraph.remove(paragraph.find('a:endParaRPr', NS))


def field_type(package):
    field_paragraph(package).find('a:fld', NS).set('type', 'datetime1')


def flattened_field(package):
    paragraph = field_paragraph(package)
    field = paragraph.find('a:fld', NS)
    run = ET.Element(tag('a', 'r'))
    run[:] = [field.find('a:rPr', NS), field.find('a:t', NS)]
    paragraph[list(paragraph).index(field)] = run


def empty_run_paragraph(package):
    run = ET.SubElement(ET.SubElement(package.shape(2, 'Bullets').find('p:txBody', NS), tag('a', 'p')), tag('a', 'r'))
    ET.SubElement(run, tag('a', 'rPr'), sz='1800', b='1')
    ET.SubElement(run, tag('a', 't'))


def resize_empty_run(package):
    package.shape(2, 'Bullets').findall('p:txBody/a:p', NS)[-1].find('a:r/a:rPr', NS).set('sz', '4800')


def shape_edit(name, path, **attributes):
    def edit(package):
        package.shape(2, name).find(path, NS).attrib.update(attributes)
    return edit


def shape_style(package):
    shape = package.shape(2, 'Bullets')
    style = ET.fromstring(f'<p:style xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>'
                          '<a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
                          '<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>')
    shape.insert(list(shape).index(shape.find('p:txBody', NS)), style)


def list_style(package):
    levels = package.shape(2, 'Bullets').find('p:txBody/a:lstStyle', NS)
    ET.SubElement(ET.SubElement(levels, tag('a', 'lvl1pPr'), marL='0'), tag('a', 'defRPr'), sz='1800')


def styled_bullets(package):
    shape_style(package)
    list_style(package)


def styled_notes(directory, name, ends, lines, keep_end=True):
    package = Package(checker.fixture(task_named('pptx-08-speaker-notes')))
    body = package.xml('ppt/notesSlides/notesSlide2.xml').find('p:cSld/p:spTree/p:sp/p:txBody', NS)
    template = body.find('a:p', NS)
    body.remove(template)
    template.insert(0, ET.Element(tag('a', 'pPr'), algn='r'))
    template.find('a:r', NS).insert(0, ET.Element(tag('a', 'rPr'), b='1', sz='1700'))
    for number, size in enumerate(ends, 1):
        paragraph = copy.deepcopy(template)
        paragraph.find('a:r/a:t', NS).text = f'Cue {number}'
        ET.SubElement(paragraph, tag('a', 'endParaRPr'), sz=size)
        body.append(paragraph)
    source = directory / f'{name}-input.pptx'
    package.save(source)
    paragraphs = body.findall('a:p', NS)
    for paragraph in paragraphs:
        body.remove(paragraph)
    for number, line in enumerate(lines):
        paragraph = copy.deepcopy(paragraphs[min(number, len(paragraphs) - 1)])
        paragraph.find('a:r/a:t', NS).text = line
        if not keep_end:
            paragraph.remove(paragraph.find('a:endParaRPr', NS))
        body.append(paragraph)
    output = directory / f'{name}-output.pptx'
    package.save(output)
    return source, output


def damage_docx(package, number):
    if number == 1:
        set_word_text(word_paragraph(package, 'Summary'), 'Summery')
    elif number == 2:
        set_word_text(word_paragraph(package, 'Beacon licences'), 'Atlas licences')
    elif number == 3:
        set_word_text(word_cell(package, 3, 4), '103')
    elif number == 4:
        for run in word_rows(package)[4].iter(tag('w', 'r')):
            bold(run)
    elif number == 5:
        set_word_text(word_cell(package, 4, 3), '620')
    elif number in (6, 8, 10):
        text, style = {6: ('Risks and mitigations', 'Heading1'), 8: ('The steering group meets on 20 October.', 'Heading1'),
                       10: ('Pilot status report', 'Title')}[number]
        word_style(word_paragraph(package, text), style)
    elif number == 7:
        runs = word_paragraph(package, PILOT).findall('w:r', NS)
        runs[1].find('w:t', NS).text = 'ten participants across'
        runs[2].find('w:t', NS).text = ' two sites.'
    elif number == 9:
        insert_word_paragraph(package, 'The table lists planned and actual costs in euros.', '')


def damage_requested_change(package, format, number):
    if format == 'docx':
        damage_docx(package, number)
    elif format == 'xlsx':
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


def validate_document_metadata(directory):
    rejected = 0
    for task_id in ('xlsx-08-organize-sheets', 'pptx-01-change-title'):
        task = next(task for task in checker.tasks() if task['id'] == task_id)
        package = Package(checker.fixture(task))
        package.xml_parts['docProps/core.xml'] = ET.fromstring(
            '<coreProperties xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">'
            '<title xmlns="http://purl.org/dc/elements/1.1/">Evaluation fixture</title></coreProperties>')
        if task['format'] == 'xlsx':
            root = package.xml('xl/workbook.xml')
            views = root.find('s:bookViews', NS)
            views[0].attrib.update(activeTab='0', firstSheet='2', xWindow='32', showSheetTabs='1')
            ET.SubElement(views, tag('s', 'workbookView'), activeTab='1', firstSheet='0', visibility='visible')
            ET.SubElement(root, tag('s', 'workbookPr'), date1904='0', filterPrivacy='0')
            root.find('s:calcPr', NS).set('calcMode', 'auto')
            names = ET.SubElement(root, tag('s', 'definedNames'))
            ET.SubElement(names, tag('s', 'definedName'), name='Revenue').text = 'Forecast!$B$2'
            ET.SubElement(names, tag('s', 'definedName'), name='Local', localSheetId='0', hidden='0').text = "'Forecast'!$A$1"
            ET.SubElement(names, tag('s', 'definedName'), name='Label').text = 'IF(1,"Forecast!B2",Forecast!B2)'
            cases = [('xl/workbook.xml', 's:bookViews/s:workbookView', 'activeTab', '0'),
                     ('xl/workbook.xml', 's:bookViews/s:workbookView', 'firstSheet', '1'),
                     ('xl/workbook.xml', 's:bookViews/s:workbookView', 'showSheetTabs', '0'),
                     ('xl/workbook.xml', 's:definedNames/s:definedName[@name="Local"]', 'localSheetId', '2'),
                     ('xl/workbook.xml', 's:definedNames/s:definedName[@name="Revenue"]', '#text', "'Actuals'!$B$2"),
                     ('xl/workbook.xml', 's:workbookPr', 'date1904', '1'),
                     ('xl/workbook.xml', 's:calcPr', 'calcMode', 'manual'),
                     ('docProps/core.xml', '{http://purl.org/dc/elements/1.1/}title', '#text', 'Changed')]
        else:
            package.xml('ppt/presentation.xml').set('showSpecialPlsOnTitleSld', '1')
            package.xml_parts['ppt/presProps.xml'] = ET.fromstring(f'<p:presentationPr xmlns:p="{NS["p"]}"><p:showPr useTimings="1"/></p:presentationPr>')
            package.xml_parts['ppt/viewProps.xml'] = ET.fromstring(f'<p:viewPr xmlns:p="{NS["p"]}" lastView="slideView"><p:normalViewPr/></p:viewPr>')
            package.xml_parts['ppt/tableStyles.xml'] = ET.fromstring(f'<a:tblStyleLst xmlns:a="{NS["a"]}" def="fixture-style"/>')
            package.xml_parts['ppt/notesMasters/notesMaster1.xml'] = ET.fromstring(f'<p:notesMaster xmlns:p="{NS["p"]}" name="Notes"/>')
            cases = [('ppt/presentation.xml', '.', 'showSpecialPlsOnTitleSld', '0'),
                     ('ppt/presentation.xml', 'p:notesSz', 'cy', '9144001'),
                     ('ppt/presProps.xml', 'p:showPr', 'useTimings', '0'),
                     ('ppt/viewProps.xml', '.', 'lastView', 'notesView'),
                     ('ppt/viewProps.xml', 'p:normalViewPr', 'showOutlineIcons', '0'),
                     ('ppt/tableStyles.xml', '.', 'def', 'changed-style'),
                     ('ppt/notesMasters/notesMaster1.xml', '.', 'name', 'Changed'),
                     ('docProps/core.xml', '{http://purl.org/dc/elements/1.1/}title', '#text', 'Changed')]
        source = directory / f'metadata-input.{task["format"]}'
        package.save(source)
        if task['format'] == 'xlsx':
            correct_xlsx(package, 8)
            views[0].attrib.update(activeTab='1', firstSheet='0')
            views[1].attrib.update(activeTab='2', firstSheet='1')
            names[0].text = 'Outlook!$B$2'
            names[1].set('localSheetId', '1')
            names[1].text = "'Outlook'!$A$1"
            names[2].text = 'IF(1,"Forecast!B2",Outlook!B2)'
        else:
            correct_pptx(package, 1)
        output = directory / f'metadata-output.{task["format"]}'
        package.save(output)
        with patch.object(checker, 'fixture', return_value=source):
            accepts(task, output)
            for i, (part, path, attribute, value) in enumerate(cases):
                changed = Package(output)
                node = changed.xml(part).find(path, NS)
                if attribute == '#text':
                    node.text = value
                else:
                    node.set(attribute, value)
                corrupted = directory / f'metadata-{i}.{task["format"]}'
                changed.save(corrupted)
                rejects(task, corrupted)
                rejected += 1
        print(f'{task_id}: enriched metadata preserved PASS; {len(cases)} metadata edits FAIL')
    assert rejected == 16


def validate(directory):
    document_failures = []
    for task in checker.tasks():
        fixture = checker.fixture(task)
        package = Package(fixture)
        number = int(task['id'].split('-')[1])
        CORRECT[task['format']](package, number)
        output = directory / f'{task["id"]}.{task["format"]}'
        package.save(output)
        accepts(task, output)
        rejects(task, fixture, UNTOUCHED)
        if task['format'] == 'docx':
            set_word_text(word_paragraph(package, 'Collect the remaining feedback forms.'), 'Unexpected edit')
        elif task['format'] == 'xlsx':
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
        if task['id'] in ('xlsx-03-replace-product', 'xlsx-08-organize-sheets', 'pptx-01-change-title'):
            package = Package(output)
            if task['format'] == 'xlsx':
                root = package.xml('xl/workbook.xml')
                views = root.find('s:bookViews', NS)
                if views is None:
                    views = ET.SubElement(root, tag('s', 'bookViews'))
                view = views.find('s:workbookView', NS)
                if view is None:
                    view = ET.SubElement(views, tag('s', 'workbookView'))
                if task['id'] == 'xlsx-08-organize-sheets':
                    view.attrib.pop('activeTab')
                    view.attrib.pop('firstSheet')
                else:
                    view.set('activeTab', str(1 - int(view.get('activeTab', '0'))))
            else:
                root = package.xml('ppt/presentation.xml')
                root.set('showSpecialPlsOnTitleSld', '0' if root.get('showSpecialPlsOnTitleSld', '1') == '1' else '1')
            changed = directory / f'{task["id"]}-document.{task["format"]}'
            package.save(changed)
            try:
                rejects(task, changed)
                print(f'{task["id"]}: document-level edit FAIL')
            except AssertionError as error:
                document_failures.append(str(error))
                print(f'{task["id"]}: document-level edit ACCEPTED')
    assert not document_failures, '\n'.join(document_failures)
    validate_document_metadata(directory)
    validate_models(directory)
    validate_docx(directory)
    print('; '.join(f'{passed}/{ran} {kind}' for kind, (passed, ran) in COUNTS.items()))
    assert not FAILURES, '\n'.join(FAILURES)


def validate_models(directory):
    xlsx03, xlsx04, xlsx08 = (task_named(name) for name in ('xlsx-03-replace-product', 'xlsx-04-style-header', 'xlsx-08-organize-sheets'))
    pptx01, pptx08, pptx10 = (task_named(name) for name in ('pptx-01-change-title', 'pptx-08-speaker-notes', 'pptx-10-insert-slide'))
    output = {task['id']: directory / f'{task["id"]}.{task["format"]}' for task in checker.tasks()}

    sheets = output[xlsx08['id']]
    expect('xlsx-08 active sheet marked selected', xlsx08, variant(directory, 'select-active', sheets, select(1, '1')), True)
    expect('xlsx-08 explicit unselected sheet', xlsx08, variant(directory, 'select-zero', sheets, select(3, '0')), True)
    expect('xlsx-08 inactive sheet selected', xlsx08, variant(directory, 'select-inactive', sheets, select(2, '1')), False)
    group = (select(1, '1'), select(2, '1'))
    grouped_source = variant(directory, 'grouped-input', directory / 'metadata-input.xlsx', *group)
    grouped = variant(directory, 'grouped-output', directory / 'metadata-output.xlsx', *group)
    expect('grouped selection kept', xlsx08, grouped, True, grouped_source)
    expect('grouped sheet unselected', xlsx08, variant(directory, 'grouped-lost', grouped, select(2, None)), False, grouped_source)

    notes = output[pptx08['id']]
    expect('pptx-08 notes part renamed', pptx08, variant(directory, 'notes-renamed', notes, rename_notes), True)
    expect('pptx-08 notes runs bold', pptx08, variant(directory, 'notes-bold', notes, bold_notes), False)
    expect('pptx-08 notes linked to another slide', pptx08, variant(directory, 'notes-owner', notes, notes_owner(2, 1)), False)
    expect('pptx-08 notes content type', pptx08, variant(directory, 'notes-type', notes, notes_content_type), False)
    expect('pptx-01 untouched notes geometry', pptx01, variant(directory, 'notes-geometry', output[pptx01['id']], notes_geometry), False)
    expect('pptx-10 inserted notes linked to another slide', pptx10,
           variant(directory, 'notes-inserted-owner', output[pptx10['id']], notes_owner(5, 2)), False)
    missing = variant(directory, 'notes-missing-input', checker.fixture(pptx08), without_notes)
    expect('pptx-08 notes created', pptx08, notes, True, missing)
    expect('pptx-08 created notes without slide link', pptx08,
           variant(directory, 'notes-created-unlinked', notes, without_notes_links), False, missing)

    header = output[xlsx04['id']]
    expect('xlsx-04 renumbered style pools', xlsx04, variant(directory, 'styles-renumbered', header, renumber_pools), True)
    expect('xlsx-04 engine-style header format', xlsx04,
           variant(directory, 'styles-engine', checker.fixture(xlsx04), engine_header), True)
    expect('xlsx-04 solid fill background', xlsx04,
           variant(directory, 'styles-solid-background', header, fill_background('solid', indexed='64')), True)
    expect('xlsx-04 header font not applied', xlsx04, variant(directory, 'styles-font-off', header, header_font_not_applied), False)
    expect('xlsx-03 quote prefix on default format', xlsx03,
           variant(directory, 'styles-quote', output[xlsx03['id']], quote_prefix), False)
    patterned_source = variant(directory, 'patterned-input', checker.fixture(xlsx03), patterned_cell)
    patterned = variant(directory, 'patterned-output', patterned_source, lambda package: correct_xlsx(package, 3))
    expect('xlsx-03 patterned fill kept', xlsx03, patterned, True, patterned_source)
    expect('xlsx-03 patterned fill background', xlsx03,
           variant(directory, 'patterned-background', patterned, fill_background('darkGray', rgb='FFFF0000')), False, patterned_source)

    for pool, attribute, entry in [('fonts', 'fontId', '<font><i/><sz val="9"/><name val="Unused"/></font>'),
                                   ('fills', 'fillId', '<fill><patternFill patternType="lightGray"/></fill>'),
                                   ('borders', 'borderId', '<border><left style="thin"/><right/><top/><bottom/><diagonal/></border>')]:
        entry = entry.replace('>', f' xmlns="{NS["s"]}">', 1)
        expect(f'xlsx-04 unused first entry in {pool}', xlsx04,
               variant(directory, f'styles-first-{pool}', header, unused_first(pool, attribute, entry)), True)
    expect('xlsx-04 custom number format renumbered', xlsx04, variant(directory, 'styles-formats', header, renumber_formats), True)
    named_source = variant(directory, 'named-input', checker.fixture(xlsx04), named_heading)
    named = variant(directory, 'named-output', named_source, lambda package: correct_xlsx(package, 4))
    reordered = variant(directory, 'font-children-reordered', named, reorder_children('xl/styles.xml', 's:fonts/s:font'))
    expect('xlsx-04 font children reordered', xlsx04, reordered, True, named_source)
    for name, child, edit in [
            ('value changed', 'sz', lambda font, node: node.set('val', '14')),
            ('child added', 'strike', lambda font, node: ET.SubElement(font, tag('s', 'strike'))),
            ('child removed', 'name', lambda font, node: font.remove(node)),
            ('child duplicated', 'sz', lambda font, node: font.append(copy.deepcopy(node))),
            ('bold child duplicated', 'b', lambda font, node: font.append(copy.deepcopy(node)))]:
        slug = name.replace(' ', '-')
        expect(f'xlsx-04 font {name}', xlsx04,
               variant(directory, f'font-{slug}', reordered, named_font_child(child, edit)), False, named_source)
    expect('xlsx-04 border children reordered', xlsx04,
           variant(directory, 'border-children-reordered', named, reorder_children('xl/styles.xml', 's:borders/s:border')),
           False, named_source)

    xlsx01 = task_named('xlsx-01-correct-value')
    for name, kind in [('implicit numeric', None), ('explicit numeric', 'n')]:
        expect(f'xlsx-01 empty formula cache ({name})', xlsx01,
               variant(directory, f'empty-cache-{kind or "implicit"}', output[xlsx01['id']], empty_cached_value(kind)),
               False, reason='Budget!B5: empty cached value')

    comment_source = variant(directory, 'comment-input', checker.fixture(xlsx03), rich_comment)
    comment = variant(directory, 'comment-output', comment_source, lambda package: correct_xlsx(package, 3))
    reordered_comment = variant(directory, 'comment-properties-reordered', comment,
                                reorder_children('xl/checker-comments.xml', './/s:rPr'))
    expect('xlsx-03 rich text properties reordered', xlsx03, reordered_comment, True, comment_source)
    expect('xlsx-03 rich text property changed', xlsx03,
           variant(directory, 'comment-property-changed', reordered_comment,
                   comment_properties(lambda props: props.find('s:sz', NS).set('val', '12'))), False, comment_source)
    expect('xlsx-03 rich text property duplicated', xlsx03,
           variant(directory, 'comment-property-duplicated', reordered_comment,
                   comment_properties(lambda props: props.append(copy.deepcopy(props.find('s:sz', NS))))), False, comment_source)

    for number, (name, edit, accept) in enumerate([
            ('named styles kept', None, True),
            ('cell style formats reordered', swap_named_formats, True),
            ('cell styles reordered', swap_cell_styles, True),
            ('cell formats reordered', move_cell_format, True),
            ('column style changed', restyle_column, False),
            ('row style changed', restyle_row, False),
            ('named style font changed', named_font, False),
            ('cell style renamed', cell_style('name', 'Heading 2'), False),
            ('cell style builtin id changed', cell_style('builtinId', '17'), False),
            ('named format attribute added', named_attribute, False),
            ('named format child added', named_child, False)]):
        result = variant(directory, f'named-{number}', named, *([edit] if edit else []))
        expect(f'xlsx-04 {name}', xlsx04, result, accept, named_source)

    expect('pptx-01 untouched notes end properties', pptx01,
           variant(directory, 'notes-end-untouched', output[pptx01['id']], end_properties('ppt/notesSlides/notesSlide3.xml')), False)
    expect('pptx-08 untouched notes end properties', pptx08,
           variant(directory, 'notes-end-other', notes, end_properties('ppt/notesSlides/notesSlide3.xml')), False)
    expect('pptx-08 replaced notes end properties', pptx08,
           variant(directory, 'notes-end-replaced', notes, end_properties('ppt/notesSlides/notesSlide2.xml')), False)
    ended_source = variant(directory, 'ended-input', checker.fixture(pptx01), ended_slides)
    ended = variant(directory, 'ended-output', ended_source, lambda package: correct_pptx(package, 1))
    for number, (name, edit, accept) in enumerate([
            ('end properties and fields kept', None, True),
            ('untouched paragraph end properties changed', changed_end, False),
            ('replaced paragraph end properties dropped', dropped_title_end, False),
            ('field type changed', field_type, False),
            ('field flattened into its run', flattened_field, False)]):
        result = variant(directory, f'ended-{number}', ended, *([edit] if edit else []))
        expect(f'pptx-01 {name}', pptx01, result, accept, ended_source)
    empty_source = variant(directory, 'empty-run-input', checker.fixture(pptx01), empty_run_paragraph)
    empty = variant(directory, 'empty-run-output', empty_source, lambda package: correct_pptx(package, 1))
    expect('pptx-01 empty formatted run kept', pptx01, empty, True, empty_source)
    expect('pptx-01 untouched empty run resized', pptx01,
           variant(directory, 'empty-run-resized', empty, resize_empty_run), False, empty_source)
    pptx04 = task_named('pptx-04-replace-term')
    field_source = variant(directory, 'term-field-input', checker.fixture(pptx04), ended_slides)
    term = variant(directory, 'term-field-output', field_source, lambda package: correct_pptx(package, 4))
    expect('pptx-04 untouched field paragraph kept', pptx04, term, True, field_source)
    expect('pptx-04 untouched field paragraph changed', pptx04,
           variant(directory, 'term-field-changed', term, field_type), False, field_source)
    lines = pptx08['expected']['notes']['2'].split('\n')
    source, result = styled_notes(directory, 'notes-styled', ['1700', '1700'], lines)
    expect('pptx-08 styled notes', pptx08, result, True, source)
    blank = dict(pptx08, expected={'notes': {'2': 'First cue\n\nThird cue'}})
    source, result = styled_notes(directory, 'notes-blank', ['1700', '1700'], blank['expected']['notes']['2'].split('\n'))
    expect('pptx-08 styled notes with an empty line', blank, result, True, source)
    source, result = styled_notes(directory, 'notes-end-dropped', ['1700', '1700'], lines, keep_end=False)
    expect('pptx-08 written notes drop end properties', pptx08, result, False, source)
    source, result = styled_notes(directory, 'notes-end-mixed', ['1700', '4800'], lines)
    expect('pptx-08 fixture notes with mixed end properties', pptx08, result, False, source, 'fixture notes')

    titled = output[pptx01['id']]
    for number, (name, edit) in enumerate([
            ('untouched shape hidden', shape_edit('Title', 'p:nvSpPr/p:cNvPr', hidden='1')),
            ('untouched shape description added', shape_edit('Title', 'p:nvSpPr/p:cNvPr', descr='Changed')),
            ('untouched shape text box flag changed', shape_edit('Bullets', 'p:nvSpPr/p:cNvSpPr', txBox='1')),
            ('untouched shape style added', shape_style),
            ('untouched shape list style changed', list_style)]):
        expect(f'pptx-01 {name}', pptx01, variant(directory, f'shape-{number}', titled, edit), False)
    styled_source = variant(directory, 'shape-styled-input', checker.fixture(pptx01), styled_bullets)
    styled = variant(directory, 'shape-styled-output', styled_source, lambda package: correct_pptx(package, 1))
    expect('pptx-01 shape style and list style kept', pptx01, styled, True, styled_source)
    expect('pptx-01 shape style changed', pptx01,
           variant(directory, 'shape-style-changed', styled, shape_edit('Bullets', 'p:style/a:fontRef', idx='major')), False, styled_source)
    expect('pptx-01 shape list style changed', pptx01,
           variant(directory, 'shape-levels-changed', styled, shape_edit('Bullets', 'p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr', sz='2000')),
           False, styled_source)


def revision(kind):
    return ET.Element(tag('w', kind), {tag('w', 'id'): '1', tag('w', 'author'): 'Reviewer', tag('w', 'date'): '2026-10-01T00:00:00Z'})


def tracked_run(text, kind):
    def edit(package):
        paragraph = word_paragraph(package, text)
        run = paragraph.find('w:r', NS)
        wrapper = revision(kind)
        if kind == 'del':
            run.find('w:t', NS).tag = tag('w', 'delText')
        paragraph[list(paragraph).index(run)] = wrapper
        wrapper.append(run)
    return edit


def tracked_format(package):
    change = revision('rPrChange')
    ET.SubElement(change, tag('w', 'rPr'))
    word_paragraph(package, PILOT).findall('w:r', NS)[1].find('w:rPr', NS).append(change)


def split_first_run(package):
    split_run(word_paragraph(package, PILOT), [('The pilot started ', None), ('in September with ', None)])


def without_word_space(index):
    def edit(package):
        text = word_paragraph(package, PILOT).findall('w:r/w:t', NS)[index]
        assert text.text and (text.text[0].isspace() or text.text[-1].isspace())
        text.attrib.pop('{http://www.w3.org/XML/1998/namespace}space', None)
    return edit


def default_word_space(package):
    text = word_paragraph(package, 'Collect the remaining feedback forms.').find('w:r/w:t', NS)
    text.text = 'alpha '
    text.set('{http://www.w3.org/XML/1998/namespace}space', 'default')


def without_default_word_space(package):
    word_paragraph(package, 'alpha ').find('w:r/w:t', NS).attrib.pop('{http://www.w3.org/XML/1998/namespace}space')


def explicit_normal(package):
    paragraph = word_paragraph(package, 'The steering group meets on 20 October.')
    word_style(paragraph, 'Normal')
    paragraph.find('w:r', NS).insert(0, ET.Element(tag('w', 'rPr')))


def unformatted_cell(package):
    run = word_cell(package, 2, 1).find('w:r', NS)
    run.remove(run.find('w:rPr', NS))


def row_after_total(package):
    table, row = word_body(package).find('w:tbl', NS), word_rows(package)[4]
    table.remove(row)
    table.append(row)


def plain_new_row(package):
    for run in word_rows(package)[4].iter(tag('w', 'r')):
        run.remove(run.find('w:rPr', NS))


def swap(parent, first, second):
    i, j = list(parent).index(first), list(parent).index(second)
    parent[i], parent[j] = second, first


def swap_list_items(package):
    swap(word_body(package), word_paragraph(package, 'Site visits are complete'), word_paragraph(package, 'Beacon training material is in review'))


def swap_heading_styles(package):
    root = package.xml('word/styles.xml')
    styles = {style.get(tag('w', 'styleId')): style for style in root.findall('w:style', NS)}
    swap(root, styles['Heading1'], styles['Heading2'])


def word_attribute(part, path, attribute, value):
    def edit(package):
        package.xml(part).find(path, NS).set(tag('w', attribute), value)
    return edit


def core_title(package):
    package.xml('docProps/core.xml').find('{http://purl.org/dc/elements/1.1/}title').text = 'Changed'


def save_time(value):
    def edit(package):
        root = package.xml('docProps/core.xml')
        node = root.find(checker.SAVE_TIME)
        if value is None:
            root.remove(node)
        else:
            node.text = value
    return edit


def direct_heading_formatting(package):
    run = word_paragraph(package, 'Summary').find('w:r', NS)
    props = ET.Element(tag('w', 'rPr'))
    fonts = {tag('w', key): 'DejaVu Sans' for key in ('ascii', 'hAnsi', 'eastAsia', 'cs')}
    props[:] = [ET.Element(tag('w', 'rFonts'), fonts), ET.Element(tag('w', 'b')), ET.Element(tag('w', 'bCs')),
                ET.Element(tag('w', 'sz'), {tag('w', 'val'): '32'}), ET.Element(tag('w', 'szCs'), {tag('w', 'val'): '32'})]
    run.insert(0, props)


def direct_cell_spacing(package):
    word_cell(package, 3, 4).find('w:pPr', NS).insert(0, ET.Element(tag('w', 'spacing'), {
        tag('w', 'after'): '0', tag('w', 'line'): '240', tag('w', 'lineRule'): 'auto'}))


def word_empty_run(package):
    run = ET.SubElement(word_paragraph(package, 'Collect the remaining feedback forms.'), tag('w', 'r'))
    ET.SubElement(ET.SubElement(run, tag('w', 'rPr')), tag('w', 'sz'), {tag('w', 'val'): '40'})


def resize_word_empty_run(package):
    word_paragraph(package, 'Collect the remaining feedback forms.').findall('w:r', NS)[-1].find('w:rPr/w:sz', NS).set(tag('w', 'val'), '48')


def drop_word_empty_run(package):
    paragraph = word_paragraph(package, 'Collect the remaining feedback forms.')
    paragraph.remove(paragraph.findall('w:r', NS)[-1])


def paragraph_identity(text, name, value):
    def edit(package):
        paragraph = word_paragraph(package, text)
        if value is None:
            paragraph.attrib.pop(tag('w14', name))
        elif value == 'fresh':
            original = dict(paragraph.attrib)
            fresh_word_ids(package, [paragraph])
            fresh = paragraph.get(tag('w14', name))
            paragraph.attrib = original
            paragraph.set(tag('w14', name), fresh)
        else:
            paragraph.set(tag('w14', name), value)
    return edit


def fresh_inserted_paragraph_ids(package):
    fresh_word_ids(package, [word_paragraph(package, 'The steering group meets on 20 October.')])


def reused_inserted_paragraph_id(package):
    value = word_paragraph(package, 'Next steps').get(tag('w14', 'paraId'))
    word_paragraph(package, 'The steering group meets on 20 October.').set(tag('w14', 'paraId'), value)


def template_row_ids(package):
    rows = word_rows(package)
    for template, paragraph in zip(rows[3].iter(tag('w', 'p')), rows[4].iter(tag('w', 'p')), strict=True):
        for name in ('paraId', 'textId'):
            paragraph.set(tag('w14', name), template.get(tag('w14', name)))


def shared_inserted_row_paraid(package):
    first, second = list(word_rows(package)[4].iter(tag('w', 'p')))[:2]
    second.set(tag('w14', 'paraId'), first.get(tag('w14', 'paraId')))


def without_inserted_row_ids(package):
    for paragraph in word_rows(package)[4].iter(tag('w', 'p')):
        for name in ('paraId', 'textId'):
            paragraph.attrib.pop(tag('w14', name))


def swap_untouched_paragraph_ids(package):
    first = word_paragraph(package, 'Collect the remaining feedback forms.')
    second = word_paragraph(package, 'Prepare the final report for the steering group.')
    for name in ('paraId', 'textId'):
        attribute = tag('w14', name)
        before, after = first.get(attribute), second.get(attribute)
        first.set(attribute, after)
        second.set(attribute, before)


def without_word_ignorable(package):
    package.xml(DOCUMENT).attrib.pop(tag('mc', 'Ignorable'))


def validate_docx(directory):
    docx = {int(task['id'].split('-')[1]): task for task in checker.tasks() if task['format'] == 'docx'}
    output = {number: directory / f'{task["id"]}.docx' for number, task in docx.items()}
    tracked = 'tracked changes are not allowed'
    for name, number, edit, accept, reason in [
            ('docx-07 bold phrase with the first run split', 7, split_first_run, True, None),
            ('docx-07 trailing whitespace preservation removed', 7, without_word_space(0), False, None),
            ('docx-07 leading whitespace preservation removed', 7, without_word_space(2), False, None),
            ('docx-08 explicit Normal style and empty run properties', 8, explicit_normal, True, None),
            ('docx-01 tracked insertion', 1, tracked_run('Summary', 'ins'), False, tracked),
            ('docx-07 tracked formatting change', 7, tracked_format, False, tracked),
            ('docx-03 untouched cell run formatting dropped', 3, unformatted_cell, False, None),
            ('docx-04 new row after the total', 4, row_after_total, False, None),
            ('docx-04 new row without run formatting', 4, plain_new_row, False, None),
            ('docx-02 list items reordered', 2, swap_list_items, False, None),
            ('docx-06 style definitions reordered', 6, swap_heading_styles, False, None),
            ('docx-01 settings changed', 1, word_attribute('word/settings.xml', 'w:defaultTabStop', 'val', '360'), False, None),
            ('docx-01 section page size changed', 1, word_attribute(DOCUMENT, 'w:body/w:sectPr/w:pgSz', 'w', '12240'), False, None),
            ('docx-01 numbering changed', 1, word_attribute('word/numbering.xml', './/w:lvlText', 'val', '-'), False, None),
            ('docx-01 table style changed', 1, word_attribute(DOCUMENT, './/w:tblStyle', 'val', 'TableNormal'), False, None),
            ('docx-01 core properties changed', 1, core_title, False, None),
            ('docx-01 save time updated', 1, save_time('2026-10-13T12:00:00Z'), True, None),
            ('docx-01 save time removed', 1, save_time(None), False, None),
            ('docx-01 style formatting written as direct run properties', 1, direct_heading_formatting, False, None),
            ('docx-03 table style spacing written into a cell paragraph', 3, direct_cell_spacing, False, None),
            ('docx-01 untouched paragraph paraId removed', 1, paragraph_identity('Collect the remaining feedback forms.', 'paraId', None), False, None),
            ('docx-01 untouched paragraph textId removed', 1, paragraph_identity('Collect the remaining feedback forms.', 'textId', None), False, None),
            ('docx-01 untouched paragraph paraId rewritten', 1, paragraph_identity('Collect the remaining feedback forms.', 'paraId', 'fresh'), False, None),
            ('docx-01 edited heading textId rewritten', 1, paragraph_identity('Summary', 'textId', 'fresh'), False, None),
            ('docx-03 edited cell paragraph paraId removed', 3, paragraph_identity('130', 'paraId', None), False, None),
            ('docx-08 inserted paragraph with fresh ids', 8, fresh_inserted_paragraph_ids, True, None),
            ('docx-08 inserted paragraph reusing an existing paraId', 8, reused_inserted_paragraph_id, False, None),
            ('docx-08 inserted paragraph paraId out of range', 8, paragraph_identity('The steering group meets on 20 October.', 'paraId', '80000000'), False, None),
            ('docx-08 inserted paragraph paraId not eight hex digits', 8, paragraph_identity('The steering group meets on 20 October.', 'paraId', '0000001'), False, None),
            ('docx-04 inserted row keeping template ids', 4, template_row_ids, False, None),
            ('docx-04 inserted row without ids', 4, without_inserted_row_ids, True, None),
            ('docx-04 inserted row cells sharing a fresh paraId', 4, shared_inserted_row_paraid, False, None),
            ('docx-08 inserted paragraph with a fresh paraId only', 8, paragraph_identity('The steering group meets on 20 October.', 'paraId', 'fresh'), True, None),
            ('docx-01 untouched paragraphs swapped ids', 1, swap_untouched_paragraph_ids, False, None),
            ('docx-01 document Ignorable removed', 1, without_word_ignorable, False, None)]:
        slug = name.split(' ', 1)[1].replace(' ', '-')
        expect(name, docx[number], variant(directory, f'word-{slug}', output[number], edit), accept, reason=reason)
    expect('docx-09 note kept as a tracked deletion', docx[9],
           variant(directory, 'word-tracked-deletion', checker.fixture(docx[9]), tracked_run(NOTE, 'del')), False, reason=tracked)
    source = variant(directory, 'word-default-space-input', checker.fixture(docx[1]), default_word_space)
    result = variant(directory, 'word-default-space-output', source, lambda package: correct_docx(package, 1), without_default_word_space)
    expect('docx-01 default whitespace attribute omitted', docx[1], result, True, source)
    source = variant(directory, 'word-empty-input', checker.fixture(docx[1]), word_empty_run)
    result = variant(directory, 'word-empty-output', source, lambda package: correct_docx(package, 1))
    expect('docx-01 empty formatted run kept', docx[1], result, True, source)
    expect('docx-01 untouched empty run resized', docx[1], variant(directory, 'word-empty-resized', result, resize_word_empty_run), False, source)
    expect('docx-01 untouched empty run dropped', docx[1], variant(directory, 'word-empty-dropped', result, drop_word_empty_run), False, source)


def v3_namespaces(data):
    import io
    return dict(item for _, item in ET.iterparse(io.BytesIO(data), events=('start-ns',)))


def v3_save(package, path):
    with ZipFile(path, 'w', compression=ZIP_DEFLATED) as archive:
        for name, data in package.parts.items():
            if name in package.xml_parts and ET.tostring(package.xml_parts[name]) != ET.tostring(ET.fromstring(data)):
                root = copy.deepcopy(package.xml_parts[name])
                payload = ET.tostring(root, encoding='unicode')
                for prefix, uri in {'w14': NS['w14'], **v3_namespaces(data)}.items():
                    if prefix and prefix != 'xml' and f'xmlns:{prefix}=' not in payload:
                        root.set('xmlns:' + prefix, uri)
                data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            archive.writestr(name, data)


def v3_paragraphs(root, prefix):
    parents = {child: parent for parent in root.iter() for child in parent}
    result = []
    for paragraph in root.iter(tag(prefix, 'p')):
        node = paragraph
        while node in parents:
            node = parents[node]
            if node.tag == tag(prefix, 'p'):
                break
        else:
            result.append(paragraph)
    return result


def v3_text(node, prefix):
    return ''.join(t.text or '' for t in node.iter(tag(prefix, 't')))


def v3_replace(paragraph, old, new, prefix):
    nodes = list(paragraph.iter(tag(prefix, 't')))
    text = v3_text(paragraph, prefix)
    count = text.count(old)
    for _ in range(count):
        start = text.index(old)
        offsets = []
        for index, node in enumerate(nodes):
            offsets.extend((index, offset) for offset in range(len(node.text or '')))
        first, left = offsets[start]
        last, right = offsets[start + len(old) - 1]
        if first == last:
            nodes[first].text = nodes[first].text[:left] + new + nodes[first].text[right + 1:]
        else:
            nodes[first].text = nodes[first].text[:left] + new
            for node in nodes[first + 1:last]:
                node.text = ''
            nodes[last].text = nodes[last].text[right + 1:]
        for node in nodes:
            value = node.text or ''
            if prefix == 'w' and value and (value[0] in ' \t\r\n' or value[-1] in ' \t\r\n'):
                node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        text = v3_text(paragraph, prefix)
    return count


def v3_accept(root):
    remove = {'del', 'moveFrom', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd',
              'rPrChange', 'pPrChange', 'sectPrChange', 'tblPrChange', 'tblPrExChange', 'tblGridChange', 'trPrChange',
              'tcPrChange', 'cellDel', 'numberingChange'}
    for parent in list(root.iter()):
        items = []
        for child in parent:
            if child.tag.startswith('{' + NS['w'] + '}') and checker.local(child.tag) in remove:
                continue
            if child.tag in (tag('w', 'ins'), tag('w', 'moveTo')):
                items.extend(list(child))
            else:
                items.append(child)
        parent[:] = items


def v3_sheet_map(package):
    links = {r.get('Id'): r.get('Target') for r in package.xml('xl/_rels/workbook.xml.rels')}
    return {sheet.get('name'): ('xl/' + links[sheet.get(tag('r', 'id'))]).replace('xl//xl/', 'xl/')
            for sheet in package.xml('xl/workbook.xml').findall('s:sheets/s:sheet', NS)}


def v3_cell(package, part, address):
    root = package.xml(part)
    data = root.find('s:sheetData', NS)
    number = int(''.join(c for c in address if c.isdigit()))
    row = next((row for row in data if int(row.get('r')) == number), None)
    if row is None:
        row = ET.Element(tag('s', 'row'), r=str(number))
        index = next((i for i, item in enumerate(data) if int(item.get('r')) > number), len(data))
        data.insert(index, row)
    found = next((node for node in row if node.get('r') == address), None)
    if found is None:
        found = ET.Element(tag('s', 'c'), r=address)
        column = checker.column_number(''.join(c for c in address if c.isalpha()))
        index = next((i for i, item in enumerate(row) if checker.column_number(''.join(c for c in item.get('r') if c.isalpha())) > column), len(row))
        row.insert(index, found)
    if 's' not in found.attrib:
        column = checker.column_number(''.join(c for c in address if c.isalpha()))
        inherited = row.get('s', '0') if row.get('customFormat') in ('1', 'true') else next((n.get('style', '0') for n in root.findall('s:cols/s:col', NS) if int(n.get('min')) <= column <= int(n.get('max'))), '0')
        if inherited != '0':
            found.set('s', inherited)
    return found


def v3_set_cell(node, value, formula=None):
    old_formula = node.find('s:f', NS)
    node[:] = [n for n in node if n.tag not in (tag('s', 'f'), tag('s', 'v'), tag('s', 'is'))]
    node.attrib.pop('t', None)
    if formula is not None:
        item = copy.deepcopy(old_formula) if old_formula is not None else ET.Element(tag('s', 'f'))
        if item.get('t') != 'shared' or item.text:
            item.text = formula
        node.append(item)
    if isinstance(value, str):
        node.set('t', 'inlineStr')
        ET.SubElement(ET.SubElement(node, tag('s', 'is')), tag('s', 't')).text = value
    else:
        ET.SubElement(node, tag('s', 'v')).text = str(value)


def v3_style(package, node, patch):
    styles = package.xml('xl/styles.xml')
    xfs = styles.find('s:cellXfs', NS)
    xf = copy.deepcopy(xfs[int(node.get('s', '0'))])
    for key, child_tag in (('bold', 'b'), ('italic', 'i')):
        if key not in patch:
            continue
        fonts = styles.find('s:fonts', NS)
        font = copy.deepcopy(fonts[int(xf.get('fontId', '0'))])
        for item in font.findall('s:' + child_tag, NS):
            font.remove(item)
        ET.SubElement(font, tag('s', child_tag), val='1' if patch[key] else '0')
        xf.set('fontId', str(len(fonts)))
        xf.set('applyFont', '1')
        fonts.append(font)
        fonts.set('count', str(len(fonts)))
    if 'fill' in patch:
        fills = styles.find('s:fills', NS)
        fill = ET.Element(tag('s', 'fill'))
        pattern = ET.SubElement(fill, tag('s', 'patternFill'), patternType=patch['fill_pattern'])
        ET.SubElement(pattern, tag('s', 'fgColor'), rgb=patch['fill'])
        xf.set('fillId', str(len(fills)))
        xf.set('applyFill', '1')
        fills.append(fill)
        fills.set('count', str(len(fills)))
    if 'number_format' in patch:
        formats = styles.find('s:numFmts', NS)
        if formats is None:
            formats = ET.Element(tag('s', 'numFmts'))
            styles.insert(0, formats)
        identity = str(max([163] + [int(n.get('numFmtId')) for n in formats]) + 1)
        ET.SubElement(formats, tag('s', 'numFmt'), numFmtId=identity, formatCode=patch['number_format'])
        formats.set('count', str(len(formats)))
        xf.set('numFmtId', identity)
        xf.set('applyNumberFormat', '1')
    node.set('s', str(len(xfs)))
    xfs.append(xf)
    xfs.set('count', str(len(xfs)))


def v3_correct(package, task):
    spec = task['expected']
    prefix = 'w' if task['format'] == 'docx' else 'a'
    v3_structural(package, spec)
    if spec.get('accept_changes'):
        for name in package.parts:
            if name.startswith('word/') and name.endswith('.xml'):
                v3_accept(package.xml(name))
    for change in spec.get('edits', []):
        matches = 0
        for paragraph in v3_paragraphs(package.xml(change['part']), prefix):
            if 'paragraph' not in change or v3_text(paragraph, prefix) == change['paragraph']:
                matches += v3_replace(paragraph, change['old'], change['new'], prefix)
        assert matches, change
    if 'slide_replace' in spec:
        change = spec['slide_replace']
        for name in package.parts:
            if name.startswith('ppt/slides/slide') and name.endswith('.xml'):
                for paragraph in v3_paragraphs(package.xml(name), 'a'):
                    v3_replace(paragraph, change['old'], change['new'], 'a')
    for change in spec.get('table_cells', []):
        table = package.xml(change['part']).findall('.//a:tbl', NS)[change['table'] - 1]
        cell = table.findall('a:tr', NS)[change['row'] - 1].findall('a:tc', NS)[change['column'] - 1]
        paragraph = cell.find('a:txBody/a:p', NS)
        assert paragraph is not None
        v3_replace(paragraph, v3_text(paragraph, 'a'), change['text'], 'a')
    for change in spec.get('delete_rows', []):
        root = package.xml(change['part'])
        parents = {child: parent for parent in root.iter() for child in parent}
        rows = [row for row in root.iter(tag('w', 'tr')) if change['contains'] in v3_text(row, 'w')]
        assert len(rows) == 1
        parents[rows[0]].remove(rows[0])
    for text, style in spec.get('styles', {}).items():
        paragraph = next(p for p in package.xml('word/document.xml').iter(tag('w', 'p')) if v3_text(p, 'w') == text)
        props = paragraph.find('w:pPr', NS)
        if props is None:
            props = ET.Element(tag('w', 'pPr'))
            paragraph.insert(0, props)
        item = props.find('w:pStyle', NS)
        if item is None:
            item = ET.Element(tag('w', 'pStyle'))
            props.insert(0, item)
        item.set(tag('w', 'val'), style)
    if spec.get('swap_last_slides'):
        listing = package.xml('ppt/presentation.xml').find('p:sldIdLst', NS)
        original = list(listing)
        listing[:] = original[:-2] + original[-2:][::-1]
    for slide, text in spec.get('notes', {}).items():
        target = v3_slides(package)[int(slide) - 1]
        rels = package.xml(checker.relpath(target))
        import posixpath
        notes = next(posixpath.normpath(posixpath.join(posixpath.dirname(target), n.get('Target'))) for n in rels if n.get('Type').endswith('/notesSlide'))
        root = package.xml(notes)
        shape = next(s for s in root.findall('p:cSld/p:spTree/p:sp', NS) if s.find('p:nvSpPr/p:nvPr/p:ph', NS) is not None and s.find('p:nvSpPr/p:nvPr/p:ph', NS).get('type') == 'body')
        body = shape.find('p:txBody', NS)
        paragraphs = body.findall('a:p', NS)
        first = paragraphs[0]
        new = ET.Element(tag('a', 'p'), first.attrib)
        props = first.find('a:pPr', NS)
        if props is not None:
            new.append(copy.deepcopy(props))
        run = ET.SubElement(new, tag('a', 'r'))
        props = first.find('a:r/a:rPr', NS)
        if props is not None:
            run.append(copy.deepcopy(props))
        ET.SubElement(run, tag('a', 't')).text = text
        end = first.find('a:endParaRPr', NS)
        if end is not None:
            new.append(copy.deepcopy(end))
        body[:] = [n for n in body if n.tag != tag('a', 'p')] + [new]
    if task['format'] != 'xlsx':
        return
    sheets = v3_sheet_map(package)
    for sheet, changes in spec.get('cell_replace', {}).items():
        strings = [v3_text(n, 's') for n in package.xml('xl/sharedStrings.xml')] if 'xl/sharedStrings.xml' in package.parts else []
        for address, change in changes.items():
            node = v3_cell(package, sheets[sheet], address)
            value = strings[int(node.findtext('s:v', namespaces=NS))] if node.get('t') == 's' else v3_text(node, 's')
            assert change['old'] in value
            v3_set_cell(node, value.replace(change['old'], change['new']))
    for sheet, changes in spec.get('cells', {}).items():
        for address, patch in changes.items():
            node = v3_cell(package, sheets[sheet], address)
            if 'value' in patch:
                v3_set_cell(node, patch['value'], patch.get('formula'))
            if 'style' in patch:
                v3_style(package, node, patch['style'])
    for old, new in spec.get('rename', {}).items():
        root = package.xml('xl/workbook.xml')
        next(s for s in root.findall('s:sheets/s:sheet', NS) if s.get('name') == old).set('name', new)
        for part in list(sheets.values()) + ['xl/workbook.xml']:
            for node in package.xml(part).iter():
                if node.tag in (tag('s', 'f'), tag('s', 'definedName')) and node.text:
                    node.text = node.text.replace("'" + old + "'!", "'" + new + "'!").replace(old + '!', "'" + new + "'!")
        if 'docProps/app.xml' in package.parts:
            for node in package.xml('docProps/app.xml').iter():
                if node.tag.endswith('}lpstr') and node.text == old:
                    node.text = new


def v3_structural(package, spec):
    if 'insert_paragraphs' in spec:
        change = spec['insert_paragraphs']
        body = package.xml('word/document.xml').find('w:body', NS)
        index = next(i for i, p in enumerate(body) if p.tag == tag('w', 'p') and v3_text(p, 'w') == change['heading']) + 1
        source = body[index]
        additions = []
        for text in change['texts']:
            paragraph = ET.Element(tag('w', 'p'))
            props = source.find('w:pPr', NS)
            if props is not None:
                paragraph.append(copy.deepcopy(props))
            run = ET.SubElement(paragraph, tag('w', 'r'))
            props = source.find('w:r/w:rPr', NS)
            if props is not None:
                run.append(copy.deepcopy(props))
            ET.SubElement(run, tag('w', 't')).text = text
            additions.append(paragraph)
        body[index:index] = additions
    if 'delete_paragraph' in spec:
        body = package.xml('word/document.xml').find('w:body', NS)
        paragraph = next(p for p in body if p.tag == tag('w', 'p') and v3_text(p, 'w') == spec['delete_paragraph'])
        body.remove(paragraph)
    if 'format_phrase' in spec:
        change = spec['format_phrase']
        paragraph = next(p for p in package.xml('word/document.xml').iter(tag('w', 'p')) if v3_text(p, 'w') == change['paragraph'])
        begin = change['paragraph'].index(change['phrase'])
        finish = begin + len(change['phrase'])
        offset = 0
        for run in list(paragraph):
            if run.tag != tag('w', 'r'):
                continue
            value = v3_text(run, 'w')
            cuts = sorted({0, len(value), max(0, min(len(value), begin - offset)), max(0, min(len(value), finish - offset))})
            if offset < finish and offset + len(value) > begin:
                index = list(paragraph).index(run)
                paragraph.remove(run)
                for left, right in zip(cuts, cuts[1:]):
                    piece = copy.deepcopy(run)
                    text = piece.find('w:t', NS)
                    text.text = value[left:right]
                    if text.text[0] in ' \t\r\n' or text.text[-1] in ' \t\r\n':
                        text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    if begin <= offset + left < finish:
                        props = piece.find('w:rPr', NS)
                        if props is None:
                            props = ET.Element(tag('w', 'rPr'))
                            piece.insert(0, props)
                        prop = props.find('w:' + change['property'], NS)
                        if prop is None:
                            prop = ET.SubElement(props, tag('w', change['property']))
                        prop.set(tag('w', 'val'), '1')
                    paragraph.insert(index, piece)
                    index += 1
            offset += len(value)
    if 'sort_range' in spec:
        change = spec['sort_range']
        sheet = v3_sheet_map(package)[change['sheet']]
        import re
        first, low, last, high = re.fullmatch(r'([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)', change['range']).groups()
        data = package.xml(sheet).find('s:sheetData', NS)
        strings = [v3_text(si, 's') for si in package.xml('xl/sharedStrings.xml')]
        rows = {int(row.get('r')): row for row in data}
        def value(row):
            node = rows[row].find(f's:c[@r="{change["key"]}{row}"]', NS)
            return (strings[int(node.findtext('s:v', namespaces=NS))] if node.get('t') == 's' else v3_text(node, 's')).casefold()
        order = sorted(range(int(low), int(high) + 1), key=value, reverse=not change['ascending'])
        snapshots = {number: copy.deepcopy(rows[number]) for number in order}
        for destination, origin in zip(range(int(low), int(high) + 1), order):
            target = rows[destination]
            for cell in list(target):
                col = re.sub('[0-9]+$', '', cell.get('r', ''))
                if checker.column_number(first) <= checker.column_number(col) <= checker.column_number(last):
                    target.remove(cell)
            for cell in snapshots[origin]:
                col = re.sub('[0-9]+$', '', cell.get('r', ''))
                if checker.column_number(first) <= checker.column_number(col) <= checker.column_number(last):
                    cell = copy.deepcopy(cell)
                    cell.set('r', col + str(destination))
                    target.append(cell)
            target[:] = sorted(target, key=lambda c: checker.column_number(re.sub('[0-9]+$', '', c.get('r'))))
    if 'move_slide' in spec:
        change = spec['move_slide']
        listing = package.xml('ppt/presentation.xml').find('p:sldIdLst', NS)
        order = list(listing)
        order.insert(change['to'] - 1, order.pop(change['from'] - 1))
        listing[:] = order
    if 'delete_slide' in spec:
        import posixpath
        index = spec['delete_slide'] - 1
        listing = package.xml('ppt/presentation.xml').find('p:sldIdLst', NS)
        part = v3_slides(package)[index]
        links = package.xml(checker.relpath(part))
        notes = [posixpath.normpath(posixpath.join(posixpath.dirname(part), n.get('Target'))).lstrip('/')
                 for n in links if n.get('Type').endswith('/notesSlide')]
        identity = listing[index].get(tag('r', 'id'))
        del listing[index]
        links = package.xml('ppt/_rels/presentation.xml.rels')
        links[:] = [n for n in links if n.get('Id') != identity]
        removed = {part, checker.relpath(part)}
        for note in notes:
            removed.update((note, checker.relpath(note)))
        for name in removed:
            package.parts.pop(name, None)
            package.xml_parts.pop(name, None)
        types = package.xml('[Content_Types].xml')
        types[:] = [n for n in types if n.get('PartName', '').lstrip('/') not in removed]


def v3_slides(package):
    import posixpath
    rels = {n.get('Id'): n.get('Target') for n in package.xml('ppt/_rels/presentation.xml.rels')}
    return [posixpath.normpath(posixpath.join('ppt', rels[n.get(tag('r', 'id'))])).lstrip('/') for n in package.xml('ppt/presentation.xml').findall('p:sldIdLst/p:sldId', NS)]


def v3_mutants(package, task, original):
    fmt = task['format']
    def unrelated_text(p):
        prefix = 'w' if fmt == 'docx' else 'a' if fmt == 'pptx' else 's'
        original_texts = set()
        for name in original.parts:
            if name.endswith('.xml') and name.startswith({'docx':'word/', 'xlsx':'xl/worksheets/', 'pptx':'ppt/slides/'}[fmt]):
                original_texts.update(n.text for n in original.xml(name).iter(tag(prefix, 't')) if n.text)
        for name in p.parts:
            if name.endswith('.xml') and name.startswith({'docx':'word/', 'xlsx':'xl/worksheets/', 'pptx':'ppt/slides/'}[fmt]):
                for n in p.xml(name).iter(tag(prefix, 't')):
                    if n.text in original_texts:
                        n.text += ' CORRUPTED'
                        return
        if fmt == 'xlsx':
            for sheet in v3_sheet_map(p).values():
                root = p.xml(sheet)
                for cell in root.findall('s:sheetData/s:row/s:c', NS):
                    address = cell.get('r')
                    source = original.xml(sheet).find(f'.//s:c[@r="{address}"]', NS)
                    if source is not None and ET.tostring(source) == ET.tostring(cell) and cell.find('s:v', NS) is not None:
                        v3_set_cell(cell, 'CORRUPTED')
                        return
        raise ValueError('no untouched text target')
    def unrelated_property(p):
        if fmt == 'xlsx':
            sheets = v3_sheet_map(p)
            for sheet, part in sheets.items():
                for node in p.xml(part).findall('s:sheetData/s:row/s:c', NS):
                    if node.get('r') not in task['expected'].get('cells', {}).get(sheet, {}):
                        v3_style(p, node, {'bold':True, 'number_format':'0.0000000000'})
                        return
        else:
            prefix = 'w' if fmt == 'docx' else 'a'
            for name in p.parts:
                if name.startswith('word/' if fmt == 'docx' else 'ppt/slides/') and name.endswith('.xml'):
                    source_runs = {ET.tostring(run) for run in original.xml(name).iter(tag(prefix, 'r'))}
                    for run in p.xml(name).iter(tag(prefix, 'r')):
                        if v3_text(run, prefix) and ET.tostring(run) in source_runs:
                            props = run.find(prefix + ':rPr', NS)
                            if props is None:
                                props = ET.Element(tag(prefix, 'rPr'))
                                run.insert(0, props)
                            if prefix == 'w':
                                ET.SubElement(props, tag('w', 'color'), {tag('w', 'val'):'123456'})
                            else:
                                props.set('sz', '12345')
                            return
    def drop_part(p):
        import posixpath
        from urllib.parse import unquote
        name = 'docProps/custom.xml' if 'docProps/custom.xml' in p.parts else 'docProps/core.xml'
        del p.parts[name]
        p.xml_parts.pop(name, None)
        for path in p.parts:
            if path.endswith('.rels'):
                folder, filename = posixpath.split(path)
                part = posixpath.join(posixpath.dirname(folder), filename[:-5]) if path != '_rels/.rels' else ''
                root = p.xml(path)
                for item in list(root):
                    target = posixpath.normpath(posixpath.join(posixpath.dirname(part), unquote(item.get('Target')))).lstrip('/')
                    if item.get('TargetMode') != 'External' and target == name:
                        root.remove(item)
        root = p.xml('[Content_Types].xml')
        for item in list(root):
            if item.get('PartName', '').lstrip('/') == name:
                root.remove(item)
    def broken_relationship(p):
        name = next(n for n in p.parts if n.endswith('.rels'))
        item = next(n for n in p.xml(name) if n.get('TargetMode') != 'External')
        item.set('Target', 'missing-v3-part.xml')
    result = [('untouched-text', unrelated_text), ('untouched-property', unrelated_property), ('dropped-part', drop_part), ('broken-relationship', broken_relationship)]
    if fmt == 'docx':
        def duplicate(p):
            paragraphs = list(p.xml('word/document.xml').iter(tag('w', 'p')))
            paragraphs[0].set(tag('w14', 'paraId'), '7ABCDE01')
            paragraphs[1].set(tag('w14', 'paraId'), '7ABCDE01')
        def tracked(p):
            paragraph = next(n for n in p.xml('word/document.xml').iter(tag('w', 'p')) if n.find('w:r', NS) is not None)
            run = paragraph.find('w:r', NS)
            index = list(paragraph).index(run)
            paragraph.remove(run)
            change = ET.Element(tag('w', 'ins'), {tag('w','id'):'98765', tag('w','author'):'Mutation'})
            change.append(run)
            paragraph.insert(index, change)
        result += [('duplicate-paragraph-id', duplicate), ('added-tracked-change', tracked)]
    if fmt == 'xlsx':
        if any(package.xml(name).find('.//s:f', NS) is not None for name in package.parts if name.startswith('xl/worksheets/') and name.endswith('.xml')):
            def formula_value(p):
                for name in p.parts:
                    if name.startswith('xl/worksheets/') and name.endswith('.xml'):
                        for cell in p.xml(name).findall('.//s:c', NS):
                            formula = cell.find('s:f', NS)
                            if formula is not None:
                                cell.remove(formula)
                                return
            result.append(('formula-replaced-by-cache', formula_value))
    if fmt == 'pptx':
        def order(p):
            listing = p.xml('ppt/presentation.xml').find('p:sldIdLst', NS)
            listing[:] = list(listing)[1:] + list(listing)[:1]
        result.append(('slide-order', order))
        if {'move_slide', 'delete_slide', 'swap_last_slides'} & task['expected'].keys():
            def title_metadata(p):
                root = p.xml('docProps/app.xml')
                titles = next(n for n in root if n.tag.endswith('}TitlesOfParts'))[0]
                titles[-1].text = 'CORRUPTED TITLE'
            result.append(('slide-title-metadata', title_metadata))
        notes = [n for n in package.parts if n.startswith('ppt/notesSlides/notesSlide') and n.endswith('.xml')]
        if notes:
            def note(p):
                for name in notes:
                    for n in p.xml(name).iter(tag('a', 't')):
                        if n.text:
                            n.text += ' CORRUPTED'
                            return
                raise ValueError('no note text target')
            result.append(('note-text', note))
    return result


def v3_calculation_variants(task):
    def full_calculation(value='1', **attributes):
        def edit(package):
            calculation = package.xml('xl/workbook.xml').find('s:calcPr', NS)
            calculation.set('fullCalcOnLoad', value)
            calculation.attrib.update(attributes)
            package.parts.pop('xl/calcChain.xml', None)
            package.xml_parts.pop('xl/calcChain.xml', None)
            links = package.xml('xl/_rels/workbook.xml.rels')
            for item in list(links):
                if item.get('Type') == NS['r'] + '/calcChain':
                    links.remove(item)
            types = package.xml('[Content_Types].xml')
            for item in list(types):
                if item.get('PartName') == '/xl/calcChain.xml':
                    types.remove(item)
        return edit

    def numeric_drift(factor, formula):
        def edit(package):
            for part in v3_sheet_map(package).values():
                for cell in package.xml(part).findall('s:sheetData/s:row/s:c', NS):
                    value = cell.find('s:v', NS)
                    if (cell.get('t', 'n') == 'n' and (cell.find('s:f', NS) is not None) == formula
                            and value is not None and value.text):
                        number = Decimal(value.text)
                        if number.is_finite() and number != 0:
                            value.text = str(number * Decimal(factor))
                            return
            raise ValueError('no nonzero numeric drift target')
        return edit

    if task['id'] == 'xlsx-v3-01-spending-label':
        return [('full-calc-on-load-dropped-chain', full_calculation(), True),
                ('full-calc-on-load-true-dropped-chain', full_calculation('true'), True),
                ('full-calc-on-load-changed-calc-mode', full_calculation(calcMode='manual'), False),
                ('full-calc-on-load-changed-calc-id', full_calculation(calcId='0'), False),
                ('full-calc-on-load-false', full_calculation('0'), False)]
    if task['id'] in ('xlsx-v3-10-borough-correction', 'xlsx-v3-26-costing-headers'):
        return [('formula-cache-roundoff', numeric_drift('1.000000000000001', True), True),
                ('formula-cache-value-change', numeric_drift('1.01', True), False),
                ('literal-number-roundoff', numeric_drift('1.000000000000001', False), False)]
    return []


def v3_minimal_replace(paragraph, old, new, prefix, drop_empty=False, insertion_side='right'):
    import re
    nodes = list(paragraph.iter(tag(prefix, 't')))
    runs = {run: v3_text(run, prefix) for run in paragraph.iter(tag(prefix, 'r'))}
    text = v3_text(paragraph, prefix)
    leading = 0
    while leading < min(len(old), len(new)) and old[leading] == new[leading]:
        leading += 1
    trailing = 0
    while trailing < min(len(old), len(new)) - leading and old[-trailing - 1] == new[-trailing - 1]:
        trailing += 1
    inserted = new[leading:len(new) - trailing]
    starts = [match.start() for match in re.finditer(re.escape(old), text)]
    for match in reversed(starts):
        start, end = match + leading, match + len(old) - trailing
        offsets = [(index, offset) for index, node in enumerate(nodes) for offset in range(len(node.text or ''))]
        if start == end:
            if not inserted:
                continue
            if start and (insertion_side == 'left' or start == len(offsets)):
                first, left = offsets[start - 1]
                left += 1
            else:
                first, left = offsets[start]
            value = nodes[first].text
            nodes[first].text = value[:left] + inserted + value[left:]
        else:
            first, left = offsets[start]
            last, right = offsets[end - 1]
            if first == last:
                value = nodes[first].text
                nodes[first].text = value[:left] + inserted + value[right + 1:]
            else:
                nodes[first].text = nodes[first].text[:left] + inserted
                for node in nodes[first + 1:last]:
                    node.text = ''
                nodes[last].text = nodes[last].text[right + 1:]
    for node in nodes:
        value = node.text or ''
        if prefix == 'w' and value and (value[0].isspace() or value[-1].isspace()):
            node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    if drop_empty:
        parents = {child: parent for parent in paragraph.iter() for child in parent}
        for run, before in runs.items():
            if before and not v3_text(run, prefix) and all(child.tag in {tag(prefix, 'rPr'), tag(prefix, 't')} for child in run):
                parents[run].remove(run)
    return len(starts)


def v3_fidelity_variants(task, original):
    spec, fmt = task['expected'], task['format']
    result = []
    text_tasks = ('pptx-v3-14-colour-headings', 'pptx-v3-11-conference-spelling', 'docx-v3-13-review-identifier')
    if task['id'] in text_tasks:
        prefix = 'w' if fmt == 'docx' else 'a'
        def text_shapes(package, drop_empty=False, mixed=False, unchanged=False):
            count = 0
            parts = dict.fromkeys(change['part'] for change in spec.get('edits', []))
            if 'slide_replace' in spec:
                parts.update(dict.fromkeys(name for name in original.parts if name.startswith('ppt/slides/slide') and name.endswith('.xml')))
            for part in parts:
                root = package.xml(part)
                parents = {child: parent for parent in root.iter() for child in parent}
                for source, target in zip(v3_paragraphs(original.xml(part), prefix), v3_paragraphs(root, prefix)):
                    text = v3_text(source, prefix)
                    changes = [change for change in spec.get('edits', []) if change['part'] == part
                               and ('paragraph' not in change or change['paragraph'] == text)]
                    if 'slide_replace' in spec and part.startswith('ppt/slides/'):
                        changes.append(spec['slide_replace'])
                    if not any(change['old'] in text for change in changes):
                        continue
                    paragraph = copy.deepcopy(source)
                    count += 1
                    if not (unchanged and count == 1):
                        for change in changes:
                            if mixed and count % 2 == 0:
                                v3_replace(paragraph, change['old'], change['new'], prefix)
                            else:
                                v3_minimal_replace(paragraph, change['old'], change['new'], prefix, drop_empty)
                    parent = parents[target]
                    index = list(parent).index(target)
                    parent.remove(target)
                    parent.insert(index, paragraph)
            assert count >= 2, 'text shape variants need multiple edited paragraphs'
        def foreign_format(package):
            text_shapes(package)
            if fmt == 'pptx':
                part = 'ppt/slides/slide7.xml' if '14-' in task['id'] else 'ppt/slides/slide1.xml'
                old, new = ('COLOuR WHEEL', 'COLOUR WHEEL') if '14-' in task['id'] else ('Otober', 'October')
            else:
                part, old, new = 'word/header1.xml', 'MRID [number]', 'MRID 12345678'
            root = package.xml(part)
            paragraph = next(p for p in v3_paragraphs(root, prefix) if new in v3_text(p, prefix))
            start = v3_text(paragraph, prefix).index(new)
            leading = 0
            while leading < min(len(old), len(new)) and old[leading] == new[leading]:
                leading += 1
            position = 0
            for run in paragraph.iter(tag(prefix, 'r')):
                value = v3_text(run, prefix)
                if position <= start + leading < position + len(value):
                    target, offset = run, start + leading - position
                    break
                position += len(value)
            else:
                raise ValueError('no changed character for formatting mutant')
            target_props = ET.tostring(target.find(prefix + ':rPr', NS)) if target.find(prefix + ':rPr', NS) is not None else None
            donor_root = package.xml('word/document.xml') if fmt == 'docx' else root
            donor = next(run for run in donor_root.iter(tag(prefix, 'r')) if run not in set(paragraph.iter(tag(prefix, 'r')))
                         and run.find(prefix + ':rPr', NS) is not None
                         and ET.tostring(run.find(prefix + ':rPr', NS)) != target_props)
            parents = {child: parent for parent in paragraph.iter() for child in parent}
            parent = parents[target]
            index = list(parent).index(target)
            assert len(target.findall(prefix + ':t', NS)) == 1
            value = v3_text(target, prefix)
            parent.remove(target)
            for value, foreign in ((value[:offset], False), (value[offset:offset + 1], True), (value[offset + 1:], False)):
                if not value:
                    continue
                run = copy.deepcopy(target)
                run.find(prefix + ':t', NS).text = value
                if prefix == 'w':
                    run.find('w:t', NS).set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                if foreign:
                    props = run.find(prefix + ':rPr', NS)
                    if props is not None:
                        run.remove(props)
                    run.insert(0, copy.deepcopy(donor.find(prefix + ':rPr', NS)))
                parent.insert(index, run)
                index += 1
        result += [('minimal-span-kept-empty-runs', text_shapes, True),
                   ('minimal-span-dropped-empty-runs', lambda p: text_shapes(p, drop_empty=True), True),
                   ('mixed-paragraph-save-shapes', lambda p: text_shapes(p, mixed=True), True),
                   ('minimal-span-foreign-format', foreign_format, False),
                   ('minimal-span-one-target-unchanged', lambda p: text_shapes(p, unchanged=True), False)]
    if task['id'] == 'pptx-v3-03-programme-grammar':
        def reachable(package):
            import posixpath
            from urllib.parse import unquote
            seen, pending = set(), ['']
            while pending:
                part = pending.pop()
                if part in seen:
                    continue
                seen.add(part)
                path = checker.relpath(part)
                if path in package.parts:
                    pending.extend(posixpath.normpath(posixpath.join(posixpath.dirname(part), unquote(link.get('Target')))).lstrip('/')
                                   for link in package.xml(path) if link.get('TargetMode') != 'External')
            return seen - {''}
        def drop_parts(package, names):
            removed = set(names) | {checker.relpath(name) for name in names}
            for name in removed:
                package.parts.pop(name, None)
                package.xml_parts.pop(name, None)
            types = package.xml('[Content_Types].xml')
            types[:] = [node for node in types if node.get('PartName', '').lstrip('/') not in removed]
        def drop_orphans(package, single=False):
            names = (reachable(original) - reachable(package)) & package.parts.keys()
            assert {'ppt/media/image22.png', 'ppt/media/image23.png'} <= names
            drop_parts(package, {'ppt/media/image22.png'} if single else names)
        def drop_reachable(package):
            names = reachable(package)
            name = next(name for name in sorted(names) if name.startswith('ppt/media/'))
            drop_parts(package, {name})
        result += [('deleted-slide-orphans-pruned', drop_orphans, True),
                   ('deleted-slide-single-orphan-pruned', lambda p: drop_orphans(p, single=True), True),
                   ('deleted-slide-reachable-part-dropped', drop_reachable, False)]
    if task['id'] == 'xlsx-v3-04-earnings-headers':
        def volatile(package):
            cell = v3_cell(package, v3_sheet_map(package)['Contents'], 'B23')
            assert cell.find('s:f', NS).get('ca') == '1' and cell.get('t') == 'str'
            cell.find('s:v', NS).text = '© Commonwealth of Australia 2026'
            return cell
        def changed_formula(package):
            volatile(package).find('s:f', NS).text += '&"changed"'
        def changed_attributes(package):
            volatile(package).find('s:f', NS).set('ca', '0')
        def changed_type(package):
            volatile(package).set('t', 'e')
        def nonvolatile(package):
            for part in v3_sheet_map(package).values():
                for cell in package.xml(part).findall('s:sheetData/s:row/s:c', NS):
                    formula, value = cell.find('s:f', NS), cell.find('s:v', NS)
                    if formula is not None and formula.get('ca') != '1' and cell.get('t', 'n') == 'n' and value is not None and Decimal(value.text):
                        value.text = str(Decimal(value.text) * Decimal('1.01'))
                        return
            raise ValueError('no nonvolatile formula cache')
        result += [('volatile-ca1-cache', volatile, True),
                   ('volatile-ca1-formula-changed', changed_formula, False),
                   ('volatile-ca1-attributes-changed', changed_attributes, False),
                   ('volatile-ca1-type-changed', changed_type, False),
                   ('nonvolatile-formula-cache-changed', nonvolatile, False)]
    if task['id'] == 'docx-v3-13-review-identifier':
        def declared_ignorable(package):
            root = package.xml('word/theme/theme1.xml')
            root.set('xmlns:w15', 'http://schemas.microsoft.com/office/word/2012/wordml')
            root.set(tag('mc', 'Ignorable'), 'w15')
        def dropped_ignorable(package):
            root = package.xml('word/document.xml')
            tokens = root.get(tag('mc', 'Ignorable')).split()
            assert tokens
            root.set(tag('mc', 'Ignorable'), ' '.join(tokens[1:]))
        def undeclared_ignorable(package):
            root = package.xml('word/document.xml')
            root.set(tag('mc', 'Ignorable'), root.get(tag('mc', 'Ignorable')) + ' undeclared')
        def nested_ignorable(package):
            undeclared_ignorable(package)
            package.xml('word/document.xml')[0].set('xmlns:undeclared', 'urn:xt2:nested-only')
        def uri_ignorable(package):
            declared_ignorable(package)
            package.xml('word/theme/theme1.xml').set(tag('mc', 'Ignorable'), 'http://schemas.microsoft.com/office/word/2012/wordml')
        result += [('declared-w15-ignorable', declared_ignorable, True),
                   ('ignorable-existing-token-dropped', dropped_ignorable, False),
                   ('ignorable-undeclared-token', undeclared_ignorable, False),
                   ('ignorable-namespace-declared-only-on-child', nested_ignorable, False),
                   ('ignorable-uri-instead-of-prefix', uri_ignorable, False)]
    return result


def v3_allowances(task):
    fmt = task['format']
    def saved_metadata(package):
        root = package.xml('docProps/core.xml')
        tags = [('http://purl.org/dc/terms/', 'modified', '2026-10-15T12:00:00Z'),
                ('http://schemas.openxmlformats.org/package/2006/metadata/core-properties', 'lastModifiedBy', 'Office reviewer')]
        for namespace, name, value in tags:
            child = root.find('{' + namespace + '}' + name)
            if child is None:
                child = ET.SubElement(root, '{' + namespace + '}' + name)
            child.text = value
        root = package.xml('docProps/app.xml')
        key = '{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Words'
        child = root.find(key)
        if child is None:
            child = ET.SubElement(root, key)
        child.text = '12345'
    def split_run(package):
        prefix = 'w' if fmt == 'docx' else 'a'
        for name in package.parts:
            if name.startswith('word/' if fmt == 'docx' else 'ppt/slides/') and name.endswith('.xml'):
                root = package.xml(name)
                parents = {child: parent for parent in root.iter() for child in parent}
                for run in root.iter(tag(prefix, 'r')):
                    text = run.find(prefix + ':t', NS)
                    if text is None or not text.text or len(text.text) < 4:
                        continue
                    value = text.text
                    midpoint = len(value) // 2
                    other = copy.deepcopy(run)
                    text.text = value[:midpoint]
                    other.find(prefix + ':t', NS).text = value[midpoint:]
                    if prefix == 'w':
                        for node in (text, other.find(prefix + ':t', NS)):
                            node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                        other.set(tag('w', 'rsidR'), '1234ABCD')
                    parent = parents[run]
                    index = list(parent).index(run)
                    parent.insert(index + 1, other)
                    if prefix == 'w':
                        parent.insert(index + 1, ET.Element(tag('w', 'proofErr'), {tag('w', 'type'):'spellStart'}))
                        parent.insert(index + 3, ET.Element(tag('w', 'proofErr'), {tag('w', 'type'):'spellEnd'}))
                    return
        raise ValueError('no split-run target')
    result = [('save-metadata', saved_metadata)]
    if fmt in ('docx', 'pptx'):
        result.append(('split-run', split_run))
    if fmt == 'docx':
        def fresh_ids(package):
            seen = {v.upper() for name in package.parts if name.startswith('word/') and name.endswith('.xml') for node in package.xml(name).iter(tag('w', 'p')) for k, v in node.attrib.items() if k in checker.PARAGRAPH_IDS}
            number = 0x12340000
            for name in package.parts:
                if name.startswith('word/') and name.endswith('.xml'):
                    root = package.xml(name)
                    for node in root.iter(tag('w', 'p')):
                        for key in checker.PARAGRAPH_IDS:
                            if key not in node.attrib:
                                while f'{number:08X}' in seen:
                                    number += 1
                                value = f'{number:08X}'
                                seen.add(value)
                                node.set(key, value)
                    key = tag('mc', 'Ignorable')
                    root.set(key, ' '.join(dict.fromkeys((root.get(key, '') + ' w14').split())))
        result.append(('fresh-missing-identities', fresh_ids))
    if fmt == 'xlsx':
        def inline_strings(package):
            strings = package.xml('xl/sharedStrings.xml').findall('s:si', NS)
            for part in v3_sheet_map(package).values():
                for cell in package.xml(part).findall('.//s:sheetData/s:row/s:c', NS):
                    if cell.get('t') == 's':
                        value = cell.find('s:v', NS)
                        item = copy.deepcopy(strings[int(value.text)])
                        item.tag = tag('s', 'is')
                        cell.remove(value)
                        cell.set('t', 'inlineStr')
                        cell.append(item)
        result.append(('inline-string-storage', inline_strings))
    if {'move_slide', 'delete_slide', 'swap_last_slides'} & task['expected'].keys():
        def slide_metadata(package):
            root = package.xml('docProps/app.xml')
            headings = next(n for n in root if n.tag.endswith('}HeadingPairs'))[0]
            titles = next(n for n in root if n.tag.endswith('}TitlesOfParts'))[0]
            offset = 0
            for index in range(0, len(headings), 2):
                count = headings[index + 1][0]
                size = int(count.text)
                if headings[index][0].text == 'Slide Titles':
                    items = list(titles)
                    group = items[offset:offset + size]
                    spec = task['expected']
                    if 'move_slide' in spec:
                        change = spec['move_slide']
                        group.insert(change['to'] - 1, group.pop(change['from'] - 1))
                    elif 'delete_slide' in spec:
                        group.pop(spec['delete_slide'] - 1)
                    else:
                        group[-2:] = group[-2:][::-1]
                    titles[:] = items[:offset] + group + items[offset + size:]
                    titles.set('size', str(len(titles)))
                    count.text = str(len(group))
                    return
                offset += size
            raise ValueError('no slide titles metadata group')
        result.append(('updated-slide-title-metadata', slide_metadata))
    return result


def v3_partial(source, task, output):
    original = Package(source)
    partial_task = copy.deepcopy(task)
    spec = partial_task['expected']
    if len(spec) > 1:
        del spec[next(reversed(spec))]
    elif 'edits' in spec:
        if len(spec['edits']) > 1:
            spec['edits'] = spec['edits'][:1]
        else:
            change = spec['edits'][0]
            change['new'] = change['new'][:max(1,len(change['new'])//2)]
    elif 'cells' in spec:
        sheet, cells = next(iter(spec['cells'].items()))
        if len(cells) > 1:
            spec['cells'] = {sheet:{next(iter(cells)):next(iter(cells.values()))}}
        else:
            change = next(iter(cells.values()))
            if 'value' in change:
                change['value'] = change['value'][:max(1,len(change['value'])//2)] if isinstance(change['value'],str) else 0
            else:
                del change['style'][next(iter(change['style']))]
    elif 'table_cells' in spec:
        if len(spec['table_cells']) > 1:
            spec['table_cells'] = spec['table_cells'][:1]
        else:
            spec['table_cells'][0]['text'] += ' partial'
    elif 'notes' in spec:
        slide, text = next(iter(spec['notes'].items()))
        spec['notes'][slide] = text[:len(text)//2]
    elif 'cell_replace' in spec:
        change = next(iter(next(iter(spec['cell_replace'].values())).values()))
        change['new'] = change['new'][:max(1,len(change['new'])//2)]
    elif 'rename' in spec:
        old = next(iter(spec['rename']))
        spec['rename'][old] = 'Partial Version'
    elif 'insert_paragraphs' in spec:
        texts = spec['insert_paragraphs']['texts']
        if len(texts) > 1:
            spec['insert_paragraphs']['texts'] = texts[:1]
        else:
            spec['insert_paragraphs']['texts'] = [texts[0][:max(1, len(texts[0]) // 2)]]
    elif 'format_phrase' in spec:
        change = spec['format_phrase']
        change['phrase'] = change['phrase'][:max(1, len(change['phrase']) // 2)]
    else:
        spec.clear()
    partial = Package(source)
    v3_correct(partial, partial_task)
    if 'format_phrase' in spec:
        def saved(package):
            with tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / ('candidate.' + task['format'])
                v3_save(package, path)
                return Package(path).parts
        reference = Package(output).parts
        if saved(partial) == reference:
            phrase = task['expected']['format_phrase']['phrase']
            paragraph = task['expected']['format_phrase']['paragraph']
            for candidate in (phrase[len(phrase) // 2:], phrase[1:], phrase[:-1]):
                if not candidate or paragraph.count(candidate) != 1:
                    continue
                spec['format_phrase']['phrase'] = candidate
                partial = Package(source)
                v3_correct(partial, partial_task)
                if saved(partial) != reference:
                    break
            else:
                raise AssertionError(f'{task["id"]} partial-edit: no format_phrase candidate differs from the reference output')
    if 'rename' in task['expected'] and len(task['expected']) == 1:
        partial = Package(source)
        old, new = next(iter(task['expected']['rename'].items()))
        next(n for n in partial.xml('xl/workbook.xml').findall('s:sheets/s:sheet', NS) if n.get('name') == old).set('name', new)
    elif 'accept_changes' in task['expected'] and len(task['expected']) == 1:
        candidates = [(root, node) for name in partial.parts if name.startswith('word/') and name.endswith('.xml')
                      for root in [partial.xml(name)] for node in root.iter() if node.tag in checker.TRACKED]
        assert candidates, 'fixture contains no Word revisions to partially accept'
        root, node = candidates[0]
        parents = {child: parent for parent in root.iter() for child in parent}
        parent = parents[node]
        index = list(parent).index(node)
        parent.remove(node)
        if node.tag in (tag('w', 'moveTo'), tag('w', 'ins')):
            parent[index:index] = list(node)
    elif 'delete_paragraph' in task['expected']:
        paragraph = next(p for p in partial.xml('word/document.xml').iter(tag('w', 'p')) if v3_text(p, 'w') == task['expected']['delete_paragraph'])
        text = next(n for n in paragraph.iter(tag('w', 't')) if n.text)
        text.text = text.text[len(text.text) // 2:]
    elif 'sort_range' in task['expected']:
        partial = Package(output)
        change = task['expected']['sort_range']
        part = v3_sheet_map(partial)[change['sheet']]
        import re
        first, low, last, high = re.fullmatch(r'([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)', change['range']).groups()
        for number in range(int(low), int(high) + 1):
            row = partial.xml(part).find(f's:sheetData/s:row[@r="{number}"]', NS)
            address = f'{last}{number}'
            row.remove(row.find(f's:c[@r="{address}"]', NS))
            row.append(copy.deepcopy(original.xml(part).find(f's:sheetData/s:row/s:c[@r="{address}"]', NS)))
    elif 'move_slide' in task['expected']:
        change = task['expected']['move_slide']
        listing = partial.xml('ppt/presentation.xml').find('p:sldIdLst', NS)
        node = listing[change['from'] - 1]
        listing.remove(node)
        if abs(change['from'] - change['to']) > 1:
            listing.insert((change['from'] + change['to']) // 2 - 1, node)
    elif 'delete_slide' in task['expected']:
        listing = partial.xml('ppt/presentation.xml').find('p:sldIdLst', NS)
        del listing[task['expected']['delete_slide'] - 1]
    elif 'delete_rows' in task['expected'] and len(task['expected']) == 1:
        change = task['expected']['delete_rows'][0]
        row = next(n for n in partial.xml(change['part']).iter(tag('w', 'tr')) if change['contains'] in v3_text(n, 'w'))
        first = next(row.iter(tag('w', 't')))
        first.text = ''
    elif 'swap_last_slides' in task['expected'] and len(task['expected']) == 1:
        listing = partial.xml('ppt/presentation.xml').find('p:sldIdLst', NS)
        key = tag('r', 'id')
        left, right = listing[-2], listing[-1]
        left_id, right_id = left.get(key), right.get(key)
        left.set(key, right_id)
        right.set(key, left_id)
    return partial


def complex_save(package, path):
    ranks = {tag(prefix, name): {tag(prefix, item): index for index, item in enumerate(order.split())}
             for (prefix, name), order in checker.COMPLEX_ORDER.items()}
    for root in package.xml_parts.values():
        for node in root.iter():
            if node.tag in ranks:
                node[:] = sorted(node, key=lambda item: ranks[node.tag].get(item.tag, 1000))
    for name, root in package.xml_parts.items():
        package.parts.setdefault(name, ET.tostring(root))
    v3_save(package, path)


def complex_new_part(package, name, root, kind):
    package.xml_parts[name] = root
    package.parts[name] = ET.tostring(root)
    ET.SubElement(package.xml('[Content_Types].xml'), '{http://schemas.openxmlformats.org/package/2006/content-types}Override',
                  PartName='/' + name, ContentType=kind)


def complex_link(package, owner, target, kind):
    import posixpath
    path = checker.relpath(owner)
    if path not in package.parts:
        package.xml_parts[path] = ET.Element(tag('rel', 'Relationships'))
        package.parts[path] = ET.tostring(package.xml_parts[path])
    links = package.xml(path)
    identity = 'rIdReference' + str(len(links) + 1)
    relation_type = NS['r'] + '/' + kind
    if kind == 'commentsExtended':
        relation_type = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended'
    ET.SubElement(links, tag('rel', 'Relationship'), Id=identity, Type=relation_type,
                  Target=posixpath.relpath(target, posixpath.dirname(owner) or '.'))
    return identity


def complex_related(package, owner, kind):
    import posixpath
    path = checker.relpath(owner)
    if path not in package.parts:
        return None
    return next((posixpath.normpath(posixpath.join(posixpath.dirname(owner), n.get('Target'))).lstrip('/')
                 for n in package.xml(path) if n.get('Type').endswith('/' + kind)), None)


def reference_word_span(paragraph, phrase):
    value = v3_text(paragraph, 'w')
    assert value.count(phrase) == 1
    begin, offset, result = value.index(phrase), 0, []
    finish = begin + len(phrase)
    for run in list(paragraph):
        if run.tag != tag('w', 'r'):
            offset += len(v3_text(run, 'w'))
            continue
        text = v3_text(run, 'w')
        cuts = sorted({0, len(text), max(0, min(len(text), begin - offset)), max(0, min(len(text), finish - offset))})
        if offset < finish and offset + len(text) > begin:
            index = list(paragraph).index(run)
            paragraph.remove(run)
            for left, right in zip(cuts, cuts[1:]):
                piece = copy.deepcopy(run)
                piece[:] = [n for n in piece if n.tag == tag('w', 'rPr')]
                ET.SubElement(piece, tag('w', 't'), {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = text[left:right]
                paragraph.insert(index, piece)
                index += 1
                if begin <= offset + left < finish:
                    result.append(piece)
        offset += len(text)
    return result


def reference_drawing_text(body, value):
    old = body.findall('a:p', NS)
    source = old[0]
    for node in old:
        body.remove(node)
    for line in value.split('\n'):
        paragraph = ET.Element(tag('a', 'p'), source.attrib)
        props = source.find('a:pPr', NS)
        if props is not None:
            paragraph.append(copy.deepcopy(props))
        run = ET.SubElement(paragraph, tag('a', 'r'))
        props = source.find('a:r/a:rPr', NS)
        if props is not None:
            run.append(copy.deepcopy(props))
        ET.SubElement(run, tag('a', 't')).text = line
        end = source.find('a:endParaRPr', NS)
        if end is not None:
            paragraph.append(copy.deepcopy(end))
        body.append(paragraph)


def reference_shape(root, name):
    for identity in root.iter(tag('p', 'cNvPr')):
        if identity.get('name') == name:
            parents = {n: p for p in root.iter() for n in p}
            return parents[parents[identity]]
    raise ValueError('missing named shape')


def reference_word_table(package, change):
    tables = package.xml(change['part']).findall('.//w:tbl', NS)
    if isinstance(change['table'], str):
        return next(n for n in tables if n.find('w:tblPr/w:tblCaption', NS) is not None
                    and n.find('w:tblPr/w:tblCaption', NS).get(tag('w', 'val')) == change['table'])
    return tables[change['table'] - 1]


def reference_comment_body(node):
    paragraphs = []
    for paragraph in node.iter(tag('w', 'p')):
        pieces = []
        for item in paragraph.iter():
            if item.tag == tag('w', 't'):
                pieces.append(item.text or '')
            elif item.tag == tag('w', 'tab'):
                pieces.append('\t')
            elif item.tag in (tag('w', 'br'), tag('w', 'cr')):
                pieces.append('\n')
        paragraphs.append(''.join(pieces))
    return '\n'.join(paragraphs)


def complex_reference(package, task):
    spec, fmt = task['expected'], task['format']
    if fmt == 'docx':
        for change in spec.get('tracked_changes', []):
            root = package.xml(change['part'])
            if change['action'] in ('accept', 'reject'):
                node = next(n for n in root.iter() if n.tag in (tag('w', 'ins'), tag('w', 'del'))
                            and ''.join(t.text or '' for t in n.iter() if t.tag in (tag('w', 't'), tag('w', 'delText'))) == change['text'])
                parent = next(p for p in root.iter() if node in list(p))
                index = list(parent).index(node)
                parent.remove(node)
                if (change['action'] == 'accept') == (node.tag == tag('w', 'ins')):
                    for text in node.iter(tag('w', 'delText')):
                        text.tag = tag('w', 't')
                    parent[index:index] = list(node)
                continue
            paragraph = next(p for p in root.iter(tag('w', 'p')) if v3_text(p, 'w') == change['paragraph'])
            insertion = change['action'] == 'insert'
            runs = reference_word_span(paragraph, change['after'] if insertion else change['text'])
            identity = str(max([0] + [int(n.get(tag('w', 'id'))) for n in root.iter() if n.tag in checker.TRACKED]) + 1)
            node = ET.Element(tag('w', 'ins' if insertion else 'del'), {tag('w', 'id'): identity, tag('w', 'author'): change['author']})
            index = list(paragraph).index(runs[-1] if insertion else runs[0]) + int(insertion)
            if insertion:
                run = copy.deepcopy(runs[-1])
                run[:] = [n for n in run if n.tag == tag('w', 'rPr')]
                ET.SubElement(run, tag('w', 't'), {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = change['text']
                node.append(run)
            else:
                for run in runs:
                    paragraph.remove(run)
                    for text in run.iter(tag('w', 't')):
                        text.tag = tag('w', 'delText')
                    node.append(run)
            paragraph.insert(index, node)
        for change in spec.get('comments', []):
            owner = DOCUMENT
            part = complex_related(package, owner, 'comments')
            if part is None:
                part = 'word/referenceComments.xml'
                complex_new_part(package, part, ET.Element(tag('w', 'comments')),
                                 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml')
                complex_link(package, owner, part, 'comments')
            comments = package.xml(part)
            action = change['action']
            parent = next((n for n in comments if reference_comment_body(n) == change.get('comment')), None)
            if action != 'resolve':
                identity = str(max([-1] + [int(n.get(tag('w', 'id'))) for n in comments]) + 1)
                node = ET.SubElement(comments, tag('w', 'comment'), {tag('w', 'id'): identity, tag('w', 'author'): change['author']})
                for line in change['text'].split('\n'):
                    paragraph = ET.SubElement(node, tag('w', 'p'))
                    ET.SubElement(ET.SubElement(paragraph, tag('w', 'r')), tag('w', 't'),
                                  {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = line
            if action == 'add':
                paragraph = next(p for p in package.xml(change['part']).iter(tag('w', 'p')) if v3_text(p, 'w') == change['paragraph'])
                runs = reference_word_span(paragraph, change['phrase'])
                start, end = list(paragraph).index(runs[0]), list(paragraph).index(runs[-1]) + 1
                paragraph.insert(end, ET.Element(tag('w', 'commentRangeEnd'), {tag('w', 'id'): identity}))
                ref = ET.Element(tag('w', 'r'))
                ET.SubElement(ref, tag('w', 'commentReference'), {tag('w', 'id'): identity})
                paragraph.insert(end + 1, ref)
                paragraph.insert(start, ET.Element(tag('w', 'commentRangeStart'), {tag('w', 'id'): identity}))
            else:
                ext = complex_related(package, owner, 'commentsExtended')
                if ext is None:
                    ext = 'word/referenceCommentsExtended.xml'
                    complex_new_part(package, ext, ET.Element(tag('w15', 'commentsEx')), 'application/vnd.ms-word.commentsExtended+xml')
                    complex_link(package, owner, ext, 'commentsExtended')
                extensions = package.xml(ext)
                for item in ([parent, node] if action == 'reply' else [parent]):
                    p = list(item.iter(tag('w', 'p')))[-1]
                    if p.get(tag('w14', 'paraId')) is None:
                        used = {n.get(tag('w14', 'paraId')) for n in comments.iter(tag('w', 'p'))}
                        number = 0x12345600
                        while f'{number:08X}' in used:
                            number += 1
                        p.set(tag('w14', 'paraId'), f'{number:08X}')
                    paragraph_id = p.get(tag('w14', 'paraId'))
                    entry = next((n for n in extensions if n.get(tag('w15', 'paraId')) == paragraph_id), None)
                    if entry is None:
                        entry = ET.SubElement(extensions, tag('w15', 'commentEx'), {tag('w15', 'paraId'): paragraph_id})
                    if action == 'resolve':
                        entry.set(tag('w15', 'done'), '1')
                    elif item is node:
                        entry.set(tag('w15', 'paraIdParent'), list(parent.iter(tag('w', 'p')))[-1].get(tag('w14', 'paraId')))
        for change in spec.get('list_items', []):
            root = package.xml(change['part'])
            anchor = next(p for p in root.iter(tag('w', 'p')) if v3_text(p, 'w') == change['after'])
            paragraph = ET.Element(tag('w', 'p'))
            paragraph.append(copy.deepcopy(anchor.find('w:pPr', NS)))
            run = ET.SubElement(paragraph, tag('w', 'r'))
            props = anchor.find('w:r/w:rPr', NS)
            if props is not None:
                run.append(copy.deepcopy(props))
            ET.SubElement(run, tag('w', 't'), {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = change['text']
            parent = next(p for p in root.iter() if anchor in list(p))
            parent.insert(list(parent).index(anchor) + 1, paragraph)
        for change in spec.get('merge_cells', []):
            table = reference_word_table(package, change)
            r1, c1, r2, c2 = change['range']
            top = table.findall('w:tr', NS)[r1 - 1].findall('w:tc', NS)[c1 - 1]
            for row_index in range(r1 - 1, r2):
                row = table.findall('w:tr', NS)[row_index]
                cells = row.findall('w:tc', NS)[c1 - 1:c2]
                first = cells[0]
                props = first.find('w:tcPr', NS)
                if props is None:
                    props = ET.Element(tag('w', 'tcPr'))
                    first.insert(0, props)
                if c2 > c1:
                    ET.SubElement(props, tag('w', 'gridSpan'), {tag('w', 'val'): str(c2 - c1 + 1)})
                    widths = [n.find('w:tcPr/w:tcW', NS) for n in cells]
                    if all(n is not None and n.get(tag('w', 'type')) == 'dxa' for n in widths):
                        widths[0].set(tag('w', 'w'), str(sum(int(n.get(tag('w', 'w'))) for n in widths)))
                if r2 > r1:
                    ET.SubElement(props, tag('w', 'vMerge'), {tag('w', 'val'): 'restart' if row_index == r1 - 1 else 'continue'})
                for cell in cells[1:]:
                    first.extend(copy.deepcopy(n) for n in cell if n.tag != tag('w', 'tcPr'))
                    row.remove(cell)
                if row_index > r1 - 1:
                    top.extend(copy.deepcopy(n) for n in first if n.tag != tag('w', 'tcPr'))
                    first[:] = [props, ET.Element(tag('w', 'p'))]
        for change in spec.get('section_setup', []):
            body = package.xml(change['part']).find('w:body', NS)
            index = next(i for i, p in enumerate(body) if p.tag == tag('w', 'p') and v3_text(p, 'w') == change['anchor'])
            index += change['position'] == 'after'
            section = next(n for block in list(body)[index:] for n in block.iter(tag('w', 'sectPr')))
            previous = copy.deepcopy(section)
            if index and body[index - 1].tag == tag('w', 'p'):
                paragraph = body[index - 1]
            else:
                paragraph = ET.Element(tag('w', 'p'))
                body.insert(index, paragraph)
            props = paragraph.find('w:pPr', NS)
            if props is None:
                props = ET.Element(tag('w', 'pPr'))
                paragraph.insert(0, props)
            props.append(previous)
            for name, attrs in [('type', {'val': 'nextPage'}), ('pgSz', {'orient': change['orientation'], 'w': change['width'], 'h': change['height']}),
                                ('pgMar', change['margins'])]:
                node = section.find('w:' + name, NS)
                if node is None:
                    node = ET.SubElement(section, tag('w', name))
                node.attrib.update({tag('w', k): str(v) for k, v in attrs.items()})
        return
    if fmt == 'xlsx':
        complex_reference_xlsx(package, spec)
    else:
        complex_reference_pptx(package, spec)


def complex_set_cell(node, value, formula=None):
    old = node.find('s:f', NS)
    attrs = dict(old.attrib) if old is not None else {}
    if attrs.get('t') == 'shared':
        attrs = {k: v for k, v in attrs.items() if k not in ('t', 'si', 'ref')}
    node[:] = [n for n in node if n.tag not in (tag('s', 'f'), tag('s', 'v'), tag('s', 'is'))]
    node.attrib.pop('t', None)
    if formula is not None:
        ET.SubElement(node, tag('s', 'f'), attrs).text = formula
    if isinstance(value, str) and formula is None:
        node.set('t', 'inlineStr')
        ET.SubElement(ET.SubElement(node, tag('s', 'is')), tag('s', 't')).text = value
    elif value is not None:
        if isinstance(value, bool):
            node.set('t', 'b')
            value = str(int(value))
        elif isinstance(value, dict):
            node.set('t', 'e')
            value = value['error']
        elif isinstance(value, str):
            node.set('t', 'str')
        ET.SubElement(node, tag('s', 'v')).text = str(value)


def reference_rename_name(formula, old, new, scope, context, shadows, sheets):
    import re
    chunks = re.split(r'("(?:[^"]|"")*")', formula)
    pattern = re.compile(r"(?<![\w.\\])(?:(?:'(?:[^']|'')+'|(?:\[[^\]]+\])?[A-Za-z_\\][\w.\\]*|\[[^\]]+\])!)?[A-Za-z_\\][\w.\\]*|\[[^\]]*\]")
    for index in range(0, len(chunks), 2):
        text = chunks[index]
        def patch_name(match):
            spelling = match.group()
            pieces = spelling.rsplit('!', 1)
            name = pieces[-1]
            if name.casefold() != old.casefold() or '[' in spelling or re.match(r'\s*[\[(]', text[match.end():]):
                return spelling
            target = pieces[0] if len(pieces) == 2 else context
            if target and target.startswith("'"):
                target = target[1:-1].replace("''", "'")
            if len(pieces) == 2 and target is not None and target.casefold() not in sheets:
                return spelling
            target = target.casefold() if target is not None else None
            use = target == scope.casefold() if scope is not None else target not in shadows
            return spelling[:-len(name)] + new if use else spelling
        chunks[index] = pattern.sub(patch_name, text)
    return ''.join(chunks)


def complex_reference_xlsx(package, spec):
    sheets = v3_sheet_map(package)
    for change in spec.get('conditional_formats', []):
        styles = package.xml('xl/styles.xml')
        pool = styles.find('s:dxfs', NS)
        if pool is None:
            pool = ET.SubElement(styles, tag('s', 'dxfs'))
        index = len(pool)
        dxf = ET.SubElement(pool, tag('s', 'dxf'))
        if 'font' in change['dxf']:
            font = ET.SubElement(dxf, tag('s', 'font'))
            for key, value in change['dxf']['font'].items():
                ET.SubElement(font, tag('s', {'bold': 'b', 'italic': 'i', 'color': 'color'}[key]),
                              **({'rgb': value} if key == 'color' else {'val': '1' if value else '0'}))
        if 'fill' in change['dxf']:
            fill = ET.SubElement(dxf, tag('s', 'fill'))
            ET.SubElement(ET.SubElement(fill, tag('s', 'patternFill'), patternType='solid'),
                          tag('s', 'fgColor'), rgb=change['dxf']['fill'])
        pool.set('count', str(len(pool)))
        root = package.xml(sheets[change['sheet']])
        priority = 1 + max([0] + [int(n.get('priority')) for n in root.findall('s:conditionalFormatting/s:cfRule', NS)])
        group = ET.SubElement(root, tag('s', 'conditionalFormatting'), sqref=change['sqref'])
        attrs = {'type': change['type'], 'priority': str(priority), 'dxfId': str(index)}
        if 'operator' in change:
            attrs['operator'] = change['operator']
        if 'stopIfTrue' in change:
            attrs['stopIfTrue'] = str(int(change['stopIfTrue']))
        node = ET.SubElement(group, tag('s', 'cfRule'), attrs)
        for formula in change['formulas']:
            ET.SubElement(node, tag('s', 'formula')).text = formula
    for change in spec.get('chart', []):
        chart = package.xml(change['part']).find('c:chart', NS)
        title = chart.find('c:title', NS)
        if title is None:
            title = ET.SubElement(chart, tag('c', 'title'))
        tx = title.find('c:tx', NS)
        if tx is None:
            tx = ET.SubElement(title, tag('c', 'tx'))
        rich = tx.find('c:rich', NS)
        if rich is not None:
            reference_drawing_text(rich, change['title'])
        else:
            tx[:] = []
            rich = ET.SubElement(tx, tag('c', 'rich'))
            ET.SubElement(rich, tag('a', 'bodyPr'))
            ET.SubElement(rich, tag('a', 'lstStyle'))
            ET.SubElement(ET.SubElement(ET.SubElement(rich, tag('a', 'p')), tag('a', 'r')), tag('a', 't')).text = change['title']
        for series, patch in zip(chart.findall('.//c:ser', NS), change['series'], strict=True):
            tx = series.find('c:tx', NS)
            if tx is None:
                tx = ET.SubElement(series, tag('c', 'tx'))
            tx[:] = []
            ET.SubElement(tx, tag('c', 'v')).text = patch['name']
            for field, key in [('cat', 'categories'), ('val', 'values')]:
                node = series.find('c:' + field, NS)
                if node is None:
                    node = series.find('c:' + ('xVal' if field == 'cat' else 'yVal'), NS)
                node[:] = []
                reference = ET.SubElement(node, tag('c', 'strRef' if node.tag == tag('c', 'cat') else 'numRef'))
                ET.SubElement(reference, tag('c', 'f')).text = patch[key]
    for change in spec.get('defined_names', []):
        root = package.xml('xl/workbook.xml')
        names = root.find('s:definedNames', NS)
        if names is None:
            names = ET.SubElement(root, tag('s', 'definedNames'))
        if change['action'] == 'add':
            attrs = {'name': change['name']}
            if 'localSheetId' in change:
                attrs['localSheetId'] = str(change['localSheetId'])
            ET.SubElement(names, tag('s', 'definedName'), attrs).text = change['formula']
        else:
            old = change['old']
            scope = str(change['localSheetId']) if 'localSheetId' in change else None
            next(n for n in names if n.get('name') == old and n.get('localSheetId') == scope).set('name', change['name'])
            sheet_names = [n.get('name') for n in root.findall('s:sheets/s:sheet', NS)]
            scope_name = sheet_names[int(scope)] if scope is not None else None
            shadows = {sheet_names[int(n.get('localSheetId'))].casefold() for n in names
                       if n.get('name', '').casefold() == old.casefold() and n.get('localSheetId') is not None}
            contexts = {part: name for name, part in sheets.items()}
            for part in package.parts:
                if not part.startswith('xl/') or not part.endswith('.xml'):
                    continue
                for node in package.xml(part).iter():
                    if checker.local(node.tag) in ('f', 'formula', 'formula1', 'formula2', 'definedName') and node.text:
                        context = contexts.get(part)
                        if node.tag == tag('s', 'definedName') and node.get('localSheetId') is not None:
                            context = sheet_names[int(node.get('localSheetId'))]
                        node.text = reference_rename_name(node.text, old, change['name'], scope_name, context, shadows,
                                                          {name.casefold() for name in sheet_names})
    for change in spec.get('autofilter_sort', []):
        import re
        sheet = sheets[change['sheet']]
        root = package.xml(sheet)
        auto = root.find('s:autoFilter', NS)
        if auto is None:
            auto = ET.SubElement(root, tag('s', 'autoFilter'))
        auto.set('ref', change['ref'])
        workbook = package.xml('xl/workbook.xml')
        index = str([n.get('name') for n in workbook.findall('s:sheets/s:sheet', NS)].index(change['sheet']))
        for name in workbook.findall('s:definedNames/s:definedName', NS):
            if name.get('name', '').casefold() == '_xlnm._filterdatabase' and name.get('localSheetId') == index and '!' in (name.text or ''):
                name.text = name.text.rsplit('!', 1)[0] + '!' + re.sub(r'([A-Z]+)(\d+)', r'$\1$\2', change['ref'])
        first, low, last, high = re.fullmatch(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', change['range']).groups()
        low, high = int(low), int(high)
        strings = [v3_text(si, 's') for si in package.xml('xl/sharedStrings.xml')] if 'xl/sharedStrings.xml' in package.parts else []
        data = root.find('s:sheetData', NS)
        rows = {int(n.get('r')): n for n in data}
        snapshots = copy.deepcopy(rows)
        masters = {n.find('s:f', NS).get('si'): (n.get('r'), n.find('s:f', NS).text)
                   for n in root.findall('.//s:c', NS) if n.find('s:f', NS) is not None
                   and n.find('s:f', NS).get('t') == 'shared' and n.find('s:f', NS).text}
        def value(row, column):
            node = rows[row].find(f's:c[@r="{column}{row}"]', NS)
            if node is None:
                return (True, 0, '')
            text = node.findtext('s:v', namespaces=NS)
            result = strings[int(text)] if node.get('t') == 's' else v3_text(node, 's') if node.get('t') == 'inlineStr' else float(text) if text else None
            return (result is None, int(isinstance(result, str)), result.casefold() if isinstance(result, str) else result if result is not None else '')
        order = list(range(low, high + 1))
        for key in change['keys'][::-1]:
            order.sort(key=lambda row: value(row, key['column']), reverse=not key['ascending'])
        for destination, origin in zip(range(low, high + 1), order):
            row = rows[destination]
            row[:] = [n for n in row if not checker.column_number(first) <= checker.column_number(re.sub(r'\d+$', '', n.get('r'))) <= checker.column_number(last)]
            for cell in snapshots[origin]:
                column = re.sub(r'\d+$', '', cell.get('r'))
                if checker.column_number(first) <= checker.column_number(column) <= checker.column_number(last):
                    cell = copy.deepcopy(cell)
                    old = cell.get('r')
                    cell.set('r', column + str(destination))
                    formula = cell.find('s:f', NS)
                    if formula is not None:
                        if formula.get('t') == 'shared':
                            master_address, master_text = masters[formula.get('si')]
                            formula.text = checker.translate_formula(master_text, master_address, old)
                            for key in ('t', 'si', 'ref'):
                                formula.attrib.pop(key, None)
                        formula.text = checker.translate_formula(formula.text, old, cell.get('r'))
                    row.append(cell)
            row[:] = sorted(row, key=lambda n: checker.column_number(re.sub(r'\d+$', '', n.get('r'))))
    for change in spec.get('dependents', []):
        entry = change['input']
        complex_set_cell(v3_cell(package, sheets[entry['sheet']], entry['cell']), entry['value'])
        for entry in change['cells']:
            complex_set_cell(v3_cell(package, sheets[entry['sheet']], entry['cell']), entry['value'], entry['formula'])


def reference_table_ids(node, table):
    import uuid
    values = {v for n in table.iter() if checker.local(n.tag) in ('rowId', 'colId', 'creationId') for v in n.attrib.values()}
    counter = max([0] + [int(n) for n in values if n.isdigit()]) + 100
    for item in node.iter():
        if checker.local(item.tag) in ('rowId', 'colId', 'creationId'):
            for key, value in list(item.attrib.items()):
                if checker.local(key) in ('val', 'id'):
                    item.set(key, str(counter) if value.isdigit() else '{' + str(uuid.UUID(int=counter)).upper() + '}')
                    counter += 1


def complex_reference_pptx(package, spec):
    import posixpath
    for change in spec.get('layout', []):
        part = v3_slides(package)[change['slide'] - 1]
        layout = next(n for n in package.parts if n.startswith('ppt/slideLayouts/') and n.endswith('.xml')
                      and package.xml(n).find('p:cSld', NS).get('name') == change['name'])
        next(n for n in package.xml(checker.relpath(part)) if n.get('Type').endswith('/slideLayout')).set('Target', posixpath.relpath(layout, posixpath.dirname(part)))
        for name, mapping in change['placeholders'].items():
            shape = reference_shape(package.xml(part), name)
            shape.find('p:nvSpPr/p:nvPr/p:ph', NS).attrib = dict(reference_shape(package.xml(layout), mapping['layout_placeholder']).find('p:nvSpPr/p:nvPr/p:ph', NS).attrib)
            reference_drawing_text(shape.find('p:txBody', NS), mapping['text'])
    for change in spec.get('group_edits', []):
        group = reference_shape(package.xml(change['part']), change['group'])
        shape = reference_shape(group, change['shape'])
        assert sum(v3_replace(p, change['old'], change['new'], 'a') for p in shape.findall('p:txBody/a:p', NS)) == 1
    for change in spec.get('table_structure', []):
        table = package.xml(change['part']).findall('.//a:tbl', NS)[change['table'] - 1]
        rows, grid = table.findall('a:tr', NS), table.find('a:tblGrid', NS)
        index = change['index'] - 1
        if change['axis'] == 'row':
            if change['action'] == 'delete':
                table.remove(rows[index])
            else:
                row = copy.deepcopy(rows[change['template'] - 1])
                reference_table_ids(row, table)
                for cell, text in zip(row.findall('a:tc', NS), change['texts'], strict=True):
                    reference_drawing_text(cell.find('a:txBody', NS), text)
                table.insert(list(table).index(rows[index]) if index < len(rows) else len(table), row)
        elif change['action'] == 'delete':
            del grid[index]
            for row in rows:
                row.remove(row.findall('a:tc', NS)[index])
        else:
            template = change['template'] - 1
            new_column = copy.deepcopy(grid[template])
            reference_table_ids(new_column, table)
            grid.insert(index, new_column)
            for row, text in zip(rows, change['texts'], strict=True):
                cells = row.findall('a:tc', NS)
                cell = copy.deepcopy(cells[template])
                reference_table_ids(cell, table)
                reference_drawing_text(cell.find('a:txBody', NS), text)
                row.insert(list(row).index(cells[index]) if index < len(cells) else len(row), cell)
    for number, change in enumerate(spec.get('duplicate_slide', []), 1):
        source = v3_slides(package)[change['slide'] - 1]
        target = f'ppt/slides/referenceDuplicate{number}.xml'
        clones = [(source, target)]
        notes = complex_related(package, source, 'notesSlide')
        if notes:
            clones.append((notes, f'ppt/notesSlides/referenceDuplicate{number}.xml'))
        comments = complex_related(package, source, 'comments')
        if comments:
            clones.append((comments, f'ppt/comments/referenceDuplicate{number}.xml'))
        mapping = dict(clones)
        for old, new in clones:
            with_types = package.xml('[Content_Types].xml')
            kind = next(n.get('ContentType') for n in with_types if n.get('PartName') == '/' + old)
            complex_new_part(package, new, copy.deepcopy(package.xml(old)), kind)
            package.parts[new] = package.parts[old]
            rels = checker.relpath(old)
            if rels in package.parts:
                root = copy.deepcopy(package.xml(rels))
                for rel in root:
                    if rel.get('TargetMode') != 'External':
                        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(old), rel.get('Target'))).lstrip('/')
                        rel.set('Target', posixpath.relpath(mapping.get(resolved, resolved), posixpath.dirname(new)))
                package.parts[checker.relpath(new)] = ET.tostring(root)
                package.xml_parts[checker.relpath(new)] = root
        identity = complex_link(package, 'ppt/presentation.xml', target, 'slide')
        listing = package.xml('ppt/presentation.xml').find('p:sldIdLst', NS)
        new_id = str(max([255] + [int(n.get('id')) for n in listing]) + 1)
        listing.insert(change['after'], ET.Element(tag('p', 'sldId'), {'id': new_id, tag('r', 'id'): identity}))
    for change in spec.get('comments', []):
        slide = v3_slides(package)[change['slide'] - 1]
        root = package.xml(complex_related(package, slide, 'comments'))
        parent = next(n for n in root if n.findtext('p:text', namespaces=NS) == change['comment'])
        authors = package.xml(complex_related(package, 'ppt/presentation.xml', 'commentAuthors'))
        author = next((n for n in authors if n.get('name') == change['author']), None)
        if author is None:
            author = ET.SubElement(authors, tag('p', 'cmAuthor'), id=str(max([0] + [int(n.get('id')) for n in authors]) + 1),
                                   name=change['author'], initials='', lastIdx='0', clrIdx='0')
        author.set('lastIdx', str(int(author.get('lastIdx')) + 1))
        node = ET.SubElement(root, tag('p', 'cm'), authorId=author.get('id'), idx=author.get('lastIdx'), dt='2026-01-01T00:00:00Z')
        ET.SubElement(node, tag('p', 'pos'), x='0', y='0')
        ET.SubElement(node, tag('p', 'text')).text = change['text']
        ext = ET.SubElement(ET.SubElement(node, tag('p', 'extLst')), tag('p', 'ext'), uri='{C676402C-5697-4E1C-873F-D02D1690AC5C}')
        ET.SubElement(ET.SubElement(ext, tag('p15', 'threadingInfo')), tag('p15', 'parentCm'),
                      authorId=parent.get('authorId'), idx=parent.get('idx'))


def complex_cases(directory):
    cases = []
    example_ids = {'docx-v3-complex-review-agreement', 'xlsx-v3-complex-review-request', 'pptx-v3-complex-review-cover'}
    examples = [task for task in checker.tasks('v3', tier='complex') if task['id'] in example_ids]
    for example in examples:
        package = Package(checker.fixture(example))
        fmt = example['format']
        task = dict(example, id=fmt + '-v3-complex-op-matrix')
        if fmt == 'docx':
            body = package.xml(DOCUMENT).find('w:body', NS)
            anchor = next(p for p in body.iter(tag('w', 'p')) if v3_text(p, 'w') == example['expected']['list_items'][0]['after'])
            numbering = package.xml('word/numbering.xml')
            num_id = anchor.find('w:pPr/w:numPr/w:numId', NS).get(tag('w', 'val'))
            abstract_id = next(n for n in numbering.findall('w:num', NS) if n.get(tag('w', 'numId')) == num_id).find('w:abstractNumId', NS).get(tag('w', 'val'))
            level = next(n for n in numbering.findall('w:abstractNum', NS) if n.get(tag('w', 'abstractNumId')) == abstract_id).find('w:lvl', NS)
            level.find('w:numFmt', NS).set(tag('w', 'val'), 'decimal')
            level.find('w:lvlText', NS).set(tag('w', 'val'), '%1.')
            fonts = level.find('w:rPr/w:rFonts', NS)
            if fonts is not None:
                fonts.attrib = {tag('w', 'ascii'): 'Calibri', tag('w', 'hAnsi'): 'Calibri'}
            style = ET.SubElement(package.xml('word/styles.xml'), tag('w', 'style'), {tag('w', 'type'): 'paragraph', tag('w', 'styleId'): 'CheckerList'})
            ET.SubElement(style, tag('w', 'name'), {tag('w', 'val'): 'Checker List'})
            ET.SubElement(style, tag('w', 'basedOn'), {tag('w', 'val'): anchor.find('w:pPr/w:pStyle', NS).get(tag('w', 'val'))})
            ET.SubElement(style, tag('w', 'pPr')).append(copy.deepcopy(anchor.find('w:pPr/w:numPr', NS)))
            def paragraph(value):
                p = ET.Element(tag('w', 'p'))
                ET.SubElement(ET.SubElement(p, tag('w', 'r')), tag('w', 't')).text = value
                body.insert(len(body) - 1, p)
                return p
            paragraph('Proposed timeline.')
            paragraph('Remove obsolete wording.')
            for index, (kind, value) in enumerate([('ins', 'Accept this insertion.'), ('del', 'Reject this deletion.'),
                                                   ('ins', 'Reject this insertion.'), ('del', 'Accept this deletion.')], 800):
                p = paragraph(value)
                run = p[0]
                p.remove(run)
                wrapper = ET.SubElement(p, tag('w', kind), {tag('w', 'id'): str(index), tag('w', 'author'): 'Original reviewer'})
                wrapper.append(run)
                if kind == 'del':
                    run[0].tag = tag('w', 'delText')
            paragraph('Section review boundary.')
            table_number = len(package.xml(DOCUMENT).findall('.//w:tbl', NS)) + 1
            table = ET.Element(tag('w', 'tbl'))
            props = ET.SubElement(table, tag('w', 'tblPr'))
            ET.SubElement(props, tag('w', 'tblCaption'), {tag('w', 'val'): 'Checker merge table'})
            grid = ET.SubElement(table, tag('w', 'tblGrid'))
            for _ in range(3):
                ET.SubElement(grid, tag('w', 'gridCol'), {tag('w', 'w'): '1000'})
            for row in range(3):
                tr = ET.SubElement(table, tag('w', 'tr'))
                for column in range(3):
                    cell = ET.SubElement(tr, tag('w', 'tc'))
                    ET.SubElement(ET.SubElement(cell, tag('w', 'tcPr')), tag('w', 'tcW'), {tag('w', 'w'): '1000', tag('w', 'type'): 'dxa'})
                    ET.SubElement(ET.SubElement(ET.SubElement(cell, tag('w', 'p')), tag('w', 'r')), tag('w', 't')).text = f'Merge {row},{column}'
            body.insert(len(body) - 1, table)
            comments = copy.deepcopy(example['expected']['comments'])
            comments += [{'action': 'reply', 'part': DOCUMENT, 'comment': comments[0]['text'], 'author': 'Host', 'text': 'Dates confirmed.'},
                         {'action': 'resolve', 'part': DOCUMENT, 'comment': comments[0]['text']}]
            task['expected'] = {
                'tracked_changes': [
                    {'action': 'insert', 'part': DOCUMENT, 'paragraph': 'Proposed timeline.', 'after': 'Proposed', 'text': ' revised', 'author': 'Review'},
                    {'action': 'delete', 'part': DOCUMENT, 'paragraph': 'Remove obsolete wording.', 'text': 'obsolete ', 'author': 'Review'},
                    *[{'action': action, 'part': DOCUMENT, 'text': text} for action, text in [
                        ('accept', 'Accept this insertion.'), ('reject', 'Reject this deletion.'),
                        ('reject', 'Reject this insertion.'), ('accept', 'Accept this deletion.')]]],
                'comments': comments,
                'list_items': copy.deepcopy(example['expected']['list_items']),
                'merge_cells': [{'part': DOCUMENT, 'table': table_number, 'range': [1, 1, 2, 2]}],
                'section_setup': [{'part': DOCUMENT, 'anchor': 'Section review boundary.', 'position': 'before',
                                   'orientation': 'landscape', 'width': 15840, 'height': 12240,
                                   'margins': {'top': 720, 'bottom': 720, 'left': 720, 'right': 720, 'header': 360, 'footer': 360, 'gutter': 0}}]}
        elif fmt == 'xlsx':
            sheet = next(iter(v3_sheet_map(package)))
            part = v3_sheet_map(package)[sheet]
            for row, values in enumerate([['Group', 'Cost', 'Label'], ['b', 2, 'second'], ['a', 3, 'third'],
                                          ['a', 1, 'first'], ['Summary', 6, 'fixed']], 100):
                for column, value in zip(['AA', 'AB', 'AC'], values):
                    node = v3_cell(package, part, column + str(row))
                    v3_set_cell(node, value)
                    if row in (101, 103):
                        v3_style(package, node, {'bold': row == 101, 'number_format': '0.000' if row == 101 else '0.0000'})
            for address, value, formula in [('AF100', 4, None), ('AG100', 8, 'AF100*2'), ('AH100', 9, 'AG100+1'),
                                             ('AI100', 42, '6*7'), ('AJ100', 4, 'OldBudget'),
                                             ('AK100', 'OldBudget', '"OldBudget"')]:
                complex_set_cell(v3_cell(package, part, address), value, formula)
            names = package.xml('xl/workbook.xml').find('s:definedNames', NS)
            if names is None:
                names = ET.SubElement(package.xml('xl/workbook.xml'), tag('s', 'definedNames'))
            ET.SubElement(names, tag('s', 'definedName'), name='OldBudget').text = f"'{sheet}'!$AF$100"
            chart_part = 'xl/charts/checkerChart.xml'
            chart = ET.Element(tag('c', 'chartSpace'))
            plot = ET.SubElement(ET.SubElement(chart, tag('c', 'chart')), tag('c', 'plotArea'))
            bar = ET.SubElement(plot, tag('c', 'barChart'))
            ET.SubElement(bar, tag('c', 'barDir'), val='col')
            ET.SubElement(bar, tag('c', 'grouping'), val='clustered')
            series = ET.SubElement(bar, tag('c', 'ser'))
            ET.SubElement(series, tag('c', 'idx'), val='0')
            ET.SubElement(series, tag('c', 'order'), val='0')
            ET.SubElement(ET.SubElement(series, tag('c', 'tx')), tag('c', 'v')).text = 'Old series'
            for field in ('cat', 'val'):
                ET.SubElement(ET.SubElement(ET.SubElement(series, tag('c', field)), tag('c', 'numRef')), tag('c', 'f')).text = f"'{sheet}'!$AB$101:$AB$102"
            for identity in ('100', '200'):
                ET.SubElement(bar, tag('c', 'axId'), val=identity)
            for axis, identity, other, position in [('catAx', '100', '200', 'b'), ('valAx', '200', '100', 'l')]:
                node = ET.SubElement(plot, tag('c', axis))
                ET.SubElement(node, tag('c', 'axId'), val=identity)
                ET.SubElement(ET.SubElement(node, tag('c', 'scaling')), tag('c', 'orientation'), val='minMax')
                ET.SubElement(node, tag('c', 'axPos'), val=position)
                ET.SubElement(node, tag('c', 'crossAx'), val=other)
                ET.SubElement(node, tag('c', 'crosses'), val='autoZero')
            complex_new_part(package, chart_part, chart, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
            drawing = 'xl/drawings/checkerDrawing.xml'
            xdr = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
            root = ET.fromstring(f'<xdr:wsDr xmlns:xdr="{xdr}" xmlns:a="{NS["a"]}" xmlns:c="{NS["c"]}" xmlns:r="{NS["r"]}"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="4000000" cy="3000000"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Checker chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="{NS["c"]}"><c:chart r:id="rIdReference1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>')
            complex_new_part(package, drawing, root, 'application/vnd.openxmlformats-officedocument.drawing+xml')
            complex_link(package, drawing, chart_part, 'chart')
            identity = complex_link(package, part, drawing, 'drawing')
            ET.SubElement(package.xml(part), tag('s', 'drawing'), {tag('r', 'id'): identity})
            task['expected'] = {
                'conditional_formats': copy.deepcopy(example['expected']['conditional_formats']),
                'chart': [{'part': chart_part, 'title': 'Review costs', 'series': [
                    {'name': 'Costs', 'categories': f"'{sheet}'!$AA$101:$AA$103", 'values': f"'{sheet}'!$AB$101:$AB$103"}]}],
                'defined_names': [{'action': 'rename', 'old': 'OldBudget', 'name': 'ReviewBudget'},
                                  {'action': 'add', 'name': 'ReviewCosts', 'formula': f"'{sheet}'!$AB$101:$AB$103"}],
                'autofilter_sort': [{'sheet': sheet, 'ref': 'AA100:AC103', 'range': 'AA101:AC103',
                                     'keys': [{'column': 'AA', 'ascending': True}, {'column': 'AB', 'ascending': True}],
                                     'fixed_rows': [100, 104]}],
                'dependents': [{'input': {'sheet': sheet, 'cell': 'AF100', 'value': 5},
                                'cells': [{'sheet': sheet, 'cell': 'AG100', 'formula': 'AF100*2', 'value': 10},
                                          {'sheet': sheet, 'cell': 'AH100', 'formula': 'AG100+1', 'value': 11},
                                          {'sheet': sheet, 'cell': 'AJ100', 'formula': 'ReviewBudget', 'value': 5}]}]}
        else:
            tree = package.xml('ppt/slides/slide1.xml').find('p:cSld/p:spTree', NS)
            group = ET.fromstring(f'<p:grpSp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><p:nvGrpSpPr><p:cNvPr id="8001" name="Checker group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="3000000" cy="1000000"/><a:chOff x="0" y="0"/><a:chExt cx="3000000" cy="1000000"/></a:xfrm></p:grpSpPr></p:grpSp>')
            shape = copy.deepcopy(reference_shape(package.xml('ppt/slides/slide1.xml'), 'Title 1'))
            shape.find('p:nvSpPr/p:cNvPr', NS).attrib.update(id='8002', name='Checker detail')
            shape.find('p:nvSpPr/p:nvPr', NS)[:] = []
            reference_drawing_text(shape.find('p:txBody', NS), 'Grouped review text.')
            group.append(shape)
            tree.append(group)
            frame = ET.fromstring(f'<p:graphicFrame xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><p:nvGraphicFramePr><p:cNvPr id="8003" name="Checker table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="10" y="20"/><a:ext cx="3000000" cy="2000000"/></p:xfrm><a:graphic><a:graphicData uri="{NS["a"]}/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="1000000"/><a:gridCol w="2000000"/></a:tblGrid></a:tbl></a:graphicData></a:graphic></p:graphicFrame>')
            table = frame.find('.//a:tbl', NS)
            for row_index in range(2):
                row = ET.SubElement(table, tag('a', 'tr'), h=str(400000 + row_index * 100000))
                for column in range(2):
                    cell = ET.SubElement(row, tag('a', 'tc'))
                    body = ET.SubElement(cell, tag('a', 'txBody'))
                    ET.SubElement(body, tag('a', 'bodyPr'))
                    ET.SubElement(body, tag('a', 'lstStyle'))
                    ET.SubElement(ET.SubElement(ET.SubElement(body, tag('a', 'p')), tag('a', 'r')), tag('a', 't')).text = f'Table {row_index},{column}'
                    ET.SubElement(ET.SubElement(ET.SubElement(cell, tag('a', 'tcPr')), tag('a', 'solidFill')), tag('a', 'srgbClr'), val='FF0000' if column else '00FF00')
            tree.append(frame)
            import base64
            package.parts['ppt/media/checker.png'] = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=')
            types = package.xml('[Content_Types].xml')
            if not any(n.get('Extension') == 'png' for n in types):
                ET.SubElement(types, '{http://schemas.openxmlformats.org/package/2006/content-types}Default', Extension='png', ContentType='image/png')
            image_id = complex_link(package, 'ppt/slides/slide1.xml', 'ppt/media/checker.png', 'image')
            tree.append(ET.fromstring(f'<p:pic xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}" xmlns:r="{NS["r"]}"><p:nvPicPr><p:cNvPr id="8004" name="Checker picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="{image_id}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'))
            complex_new_part(package, 'ppt/checkerAuthors.xml', ET.fromstring(f'<p:cmAuthorLst xmlns:p="{NS["p"]}"><p:cmAuthor id="0" name="Original" initials="O" lastIdx="1" clrIdx="0"/></p:cmAuthorLst>'),
                             'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml')
            complex_link(package, 'ppt/presentation.xml', 'ppt/checkerAuthors.xml', 'commentAuthors')
            complex_new_part(package, 'ppt/comments/checkerComment.xml', ET.fromstring(f'<p:cmLst xmlns:p="{NS["p"]}"><p:cm authorId="0" idx="1" dt="2026-01-01T00:00:00Z"><p:pos x="10" y="20"/><p:text>Review this slide.</p:text></p:cm></p:cmLst>'),
                             'application/vnd.openxmlformats-officedocument.presentationml.comments+xml')
            complex_link(package, 'ppt/slides/slide1.xml', 'ppt/comments/checkerComment.xml', 'comments')
            task['expected'] = {
                'layout': copy.deepcopy(example['expected']['layout']),
                'group_edits': [{'part': 'ppt/slides/slide1.xml', 'group': 'Checker group', 'shape': 'Checker detail', 'old': 'review', 'new': 'approved'}],
                'table_structure': [
                    {'part': 'ppt/slides/slide1.xml', 'table': 1, 'axis': 'row', 'action': 'insert', 'index': 2, 'template': 1, 'texts': ['Inserted A', 'Inserted B']},
                    {'part': 'ppt/slides/slide1.xml', 'table': 1, 'axis': 'column', 'action': 'insert', 'index': 2, 'template': 1, 'texts': ['Column A', 'Column B', 'Column C']},
                    {'part': 'ppt/slides/slide1.xml', 'table': 1, 'axis': 'row', 'action': 'delete', 'index': 3},
                    {'part': 'ppt/slides/slide1.xml', 'table': 1, 'axis': 'column', 'action': 'delete', 'index': 3}],
                'duplicate_slide': [{'slide': 1, 'after': 1}],
                'comments': [{'slide': 1, 'comment': 'Review this slide.', 'author': 'Review', 'text': 'Reviewed and approved.'}]}
        source = directory / (task['id'] + '-input.' + fmt)
        complex_save(package, source)
        cases.append((task, source))
        if fmt == 'pptx':
            for index, change in enumerate(task['expected']['table_structure']):
                isolated = copy.deepcopy(change)
                if isolated['axis'] == 'column' and isolated['action'] == 'insert':
                    isolated['texts'] = isolated['texts'][:2]
                if isolated['action'] == 'delete':
                    isolated['index'] = 2
                cases.append((dict(task, id=f'pptx-v3-complex-table-{index}', expected={'table_structure': [isolated]}), source))
    template = next(task for task in examples if task['format'] == 'docx')
    existing_task = dict(template, id='docx-v3-complex-existing-comment', fixture='h-us-epa-117c5df2c97a.docx')
    existing_source = checker.fixture(existing_task)
    package = Package(existing_source)
    comments = package.xml(complex_related(package, DOCUMENT, 'comments'))
    texts = [reference_comment_body(n) for n in comments]
    parent_text = next(text for text in texts if text and texts.count(text) == 1)
    existing_task['expected'] = {'comments': [
        {'action': 'reply', 'part': DOCUMENT, 'comment': parent_text, 'author': 'Review', 'text': 'Reviewed for the final agreement.'},
        {'action': 'resolve', 'part': DOCUMENT, 'comment': parent_text}]}
    cases.append((existing_task, existing_source))
    return cases


def complex_mutants(task, original):
    spec, fmt = task['expected'], task['format']
    result = []
    def add(name, edit):
        result.append((name, edit))
    def content_type_missing(p):
        node = next(n for n in p.xml('[Content_Types].xml') if n.get('PartName') and n.get('PartName') != '/docProps/core.xml')
        extension = node.get('PartName').rsplit('.', 1)[-1]
        root = p.xml('[Content_Types].xml')
        root[:] = [n for n in root if n is not node and n.get('Extension') != extension]
    add('integrity-missing-content-type', content_type_missing)
    def duplicate_relation(p):
        path = next(n for n in p.parts if n.endswith('.rels'))
        root = p.xml(path)
        root.append(copy.deepcopy(root[0]))
    add('integrity-duplicate-relationship-id', duplicate_relation)
    for index, change in enumerate(spec.get('tracked_changes', [])):
        def damage(p, change=change):
            root = p.xml(change['part'])
            action = change['action']
            if action in ('insert', 'delete'):
                node = next(n for n in root.iter(tag('w', 'ins' if action == 'insert' else 'del'))
                            if n.get(tag('w', 'author')) == change['author']
                            and ''.join(t.text or '' for t in n.iter() if t.tag in (tag('w', 't'), tag('w', 'delText'))) == change['text'])
                parent = next(n for n in root.iter() if node in list(n))
                at = list(parent).index(node)
                parent.remove(node)
                if action == 'insert':
                    parent[at:at] = list(node)
            else:
                source = original.xml(change['part'])
                old = next(n for n in source.iter() if n.tag in (tag('w', 'ins'), tag('w', 'del'))
                           and ''.join(t.text or '' for t in n.iter() if t.tag in (tag('w', 't'), tag('w', 'delText'))) == change['text'])
                paragraph = next(n for n in source.iter(tag('w', 'p')) if old in list(n))
                root.find('w:body', NS).insert(0, copy.deepcopy(paragraph))
        add(f'tracked-{change["action"]}-{index}-not-a-revision', damage)
        if change['action'] in ('insert', 'delete'):
            def author(p, change=change):
                node = next(n for n in p.xml(change['part']).iter(tag('w', 'ins' if change['action'] == 'insert' else 'del'))
                            if n.get(tag('w', 'author')) == change['author'])
                node.set(tag('w', 'author'), 'Wrong reviewer')
            add(f'tracked-{index}-wrong-author', author)
    if fmt == 'docx':
        for index, change in enumerate(spec.get('comments', [])):
            def damage(p, change=change):
                part = complex_related(p, DOCUMENT, 'comments')
                if change['action'] == 'resolve':
                    comment = next(n for n in p.xml(part) if reference_comment_body(n) == change['comment'])
                    identity = list(comment.iter(tag('w', 'p')))[-1].get(tag('w14', 'paraId'))
                    ext = p.xml(complex_related(p, DOCUMENT, 'commentsExtended'))
                    next(n for n in ext if n.get(tag('w15', 'paraId')) == identity).set(tag('w15', 'done'), '0')
                elif change['action'] == 'reply':
                    comment = next(n for n in p.xml(part) if reference_comment_body(n) == change['text'])
                    identity = list(comment.iter(tag('w', 'p')))[-1].get(tag('w14', 'paraId'))
                    ext = p.xml(complex_related(p, DOCUMENT, 'commentsExtended'))
                    node = next(n for n in ext if n.get(tag('w15', 'paraId')) == identity)
                    del node.attrib[tag('w15', 'paraIdParent')]
                else:
                    node = next(n for n in p.xml(part) if reference_comment_body(n) == change['text'])
                    node.find('.//w:t', NS).text += ' WRONG'
            add(f'comments-{index}-{change["action"]}', damage)
            if change['action'] == 'add':
                def anchor(p, change=change):
                    paragraph = next(n for n in p.xml(change['part']).iter(tag('w', 'p')) if v3_text(n, 'w') == change['paragraph'])
                    marker = paragraph.find('w:commentRangeStart', NS)
                    paragraph.remove(marker)
                    paragraph.append(marker)
                add(f'comments-{index}-wrong-anchor', anchor)
        for index, change in enumerate(spec.get('list_items', [])):
            def typed(p, change=change):
                paragraph = next(n for n in p.xml(change['part']).iter(tag('w', 'p')) if v3_text(n, 'w') == change['text'])
                props = paragraph.find('w:pPr', NS)
                props.remove(props.find('w:numPr', NS))
                paragraph.find('w:r/w:t', NS).text = '3. ' + change['text']
            add(f'list-{index}-typed-number', typed)
            def level(p, change=change):
                paragraph = next(n for n in p.xml(change['part']).iter(tag('w', 'p')) if v3_text(n, 'w') == change['text'])
                paragraph.find('w:pPr/w:numPr/w:ilvl', NS).set(tag('w', 'val'), '8')
            add(f'list-{index}-wrong-level', level)
        for index, change in enumerate(spec.get('merge_cells', [])):
            for kind in ('gridSpan', 'vMerge'):
                def damage(p, change=change, kind=kind):
                    table = reference_word_table(p, change)
                    node = next(table.iter(tag('w', kind)))
                    next(n for n in table.iter() if node in list(n)).remove(node)
                if (kind == 'gridSpan' and change['range'][1] != change['range'][3]) or (kind == 'vMerge' and change['range'][0] != change['range'][2]):
                    add(f'merge-{index}-missing-{kind}', damage)
        for index, change in enumerate(spec.get('section_setup', [])):
            def orientation(p, change=change):
                sections = list(p.xml(change['part']).iter(tag('w', 'sectPr')))
                node = next(n for section in sections for n in section.findall('w:pgSz', NS) if n.get(tag('w', 'orient')) == change['orientation'])
                node.set(tag('w', 'orient'), 'portrait' if change['orientation'] == 'landscape' else 'landscape')
            add(f'section-{index}-wrong-orientation', orientation)
            def margins(p, change=change):
                next(p.xml(change['part']).iter(tag('w', 'pgMar'))).set(tag('w', 'left'), '1')
            add(f'section-{index}-damaged-other-section', margins)
    if fmt == 'xlsx':
        for index, change in enumerate(spec.get('conditional_formats', [])):
            def wrong_range(p, change=change):
                root = p.xml(v3_sheet_map(p)[change['sheet']])
                next(n for n in root.findall('s:conditionalFormatting', NS) if n.get('sqref') == change['sqref']).set('sqref', 'XFD1')
            add(f'conditional-{index}-wrong-range', wrong_range)
            def wrong_dxf(p):
                pool = p.xml('xl/styles.xml').find('s:dxfs', NS)
                node = pool[-1].find('s:fill/s:patternFill/s:fgColor', NS)
                if node is not None:
                    node.set('rgb', 'FF000000')
                else:
                    ET.SubElement(pool[-1], tag('s', 'font'))
            add(f'conditional-{index}-wrong-format', wrong_dxf)
            def wrong_formula(p, change=change):
                root = p.xml(v3_sheet_map(p)[change['sheet']])
                rule = next(n for n in root.findall('s:conditionalFormatting', NS) if n.get('sqref') == change['sqref']).find('s:cfRule', NS)
                formula = rule.find('s:formula', NS)
                if formula is None:
                    formula = ET.SubElement(rule, tag('s', 'formula'))
                formula.text = '999999'
            add(f'conditional-{index}-wrong-formula', wrong_formula)
        for index, change in enumerate(spec.get('chart', [])):
            def wrong_range(p, change=change):
                series = p.xml(change['part']).find('.//c:ser', NS)
                before = original.xml(change['part']).find('.//c:ser', NS)
                field = 'cat' if series.find('c:cat', NS) is not None else 'xVal'
                node = series.find('c:' + field + '//c:f', NS)
                node.text = before.find('c:' + field + '//c:f', NS).text
                wanted = checker.chart_range_state(change['series'][0]['categories'])
                if checker.chart_range_state(node.text) == wanted:
                    node.text = node.text.rsplit('!', 1)[0] + '!$XFD$1:$XFD$2'
                assert checker.chart_range_state(node.text) != wanted
            add(f'chart-{index}-wrong-range', wrong_range)
            def wrong_name(p, change=change):
                p.xml(change['part']).find('.//c:ser/c:tx/c:v', NS).text = 'Wrong series'
            add(f'chart-{index}-wrong-series-name', wrong_name)
            def wrong_title(p, change=change):
                p.xml(change['part']).find('.//c:title//a:t', NS).text = 'Wrong chart title'
            add(f'chart-{index}-wrong-title', wrong_title)
        for index, change in enumerate(spec.get('defined_names', [])):
            def wrong_name(p, change=change):
                node = next(n for n in p.xml('xl/workbook.xml').find('s:definedNames', NS) if n.get('name') == change['name'])
                node.set('name', 'WrongName')
            add(f'name-{index}-wrong-name', wrong_name)
            if change['action'] == 'rename':
                def stale(p, change=change):
                    for part in p.parts:
                        if not part.startswith('xl/') or not part.endswith('.xml'):
                            continue
                        for n in p.xml(part).iter():
                            if checker.local(n.tag) in ('f', 'formula', 'formula1', 'formula2', 'definedName') and n.text and change['name'] in n.text:
                                n.text = n.text.replace(change['name'], change['old'])
                                return
                    next(n for n in p.xml('xl/workbook.xml').find('s:definedNames', NS) if n.get('name') == change['name']).set('name', change['old'])
                add(f'name-{index}-stale-use', stale)
        for index, change in enumerate(spec.get('autofilter_sort', [])):
            def styles_left(p, change=change):
                part = v3_sheet_map(p)[change['sheet']]
                c1, r1, c2, r2 = checker.range_coordinates(change['range'])
                seen_styles = set()
                for row in range(r1, r2 + 1):
                    for column in range(c1, c2 + 1):
                        address = checker.column_letters(column) + str(row)
                        node = v3_cell(p, part, address)
                        before = v3_cell(original, part, address)
                        seen_styles.add(before.get('s', '0'))
                        node.attrib.pop('s', None)
                        if 's' in before.attrib:
                            node.set('s', before.get('s'))
                if len(seen_styles) == 1:
                    v3_style(p, v3_cell(p, part, checker.column_letters(c1) + str(r1)), {'number_format': '0.0000000001'})
            add(f'sort-{index}-styles-left-behind', styles_left)
            def fixed(p, change=change):
                part = v3_sheet_map(p)[change['sheet']]
                column = checker.column_letters(checker.range_coordinates(change['range'])[0])
                row = change['fixed_rows'][-1] if change['fixed_rows'] else checker.range_coordinates(change['range'])[3] + 1
                v3_set_cell(v3_cell(p, part, column + str(row)), 'Changed fixed row')
            add(f'sort-{index}-fixed-row-changed', fixed)
            def filter_ref(p, change=change):
                p.xml(v3_sheet_map(p)[change['sheet']]).find('s:autoFilter', NS).set('ref', 'A1:A2')
            add(f'sort-{index}-wrong-filter', filter_ref)
        for index, change in enumerate(spec.get('dependents', [])):
            for entry in change['cells']:
                def stale(p, entry=entry):
                    node = v3_cell(p, v3_sheet_map(p)[entry['sheet']], entry['cell'])
                    value = node.find('s:v', NS)
                    if value is None:
                        value = ET.SubElement(node, tag('s', 'v'))
                    value.text = '1' if value.text == '0' else '0'
                add(f'dependents-{index}-{entry["cell"]}-stale-cache', stale)
                def lost_formula(p, entry=entry):
                    node = v3_cell(p, v3_sheet_map(p)[entry['sheet']], entry['cell'])
                    node.remove(node.find('s:f', NS))
                add(f'dependents-{index}-{entry["cell"]}-lost-formula', lost_formula)
            def other_cache(p, change=change):
                allowed = {(n['sheet'], n['cell']) for n in change['cells']}
                for name, part in v3_sheet_map(p).items():
                    for n in p.xml(part).findall('.//s:c', NS):
                        if n.find('s:f', NS) is not None and (name, n.get('r')) not in allowed:
                            n.find('s:v', NS).text = '999999'
                            return
                raise ValueError('dependent matrix requires an unrelated formula')
            allowed = {(n['sheet'], n['cell']) for n in change['cells']}
            if any(n.find('s:f', NS) is not None and (name, n.get('r')) not in allowed
                   for name, part in v3_sheet_map(original).items() for n in original.xml(part).findall('.//s:c', NS)):
                add(f'dependents-{index}-unlisted-cache-changed', other_cache)
    if fmt == 'pptx':
        for index, change in enumerate(spec.get('layout', [])):
            def old_layout(p, change=change):
                part = v3_slides(p)[change['slide'] - 1]
                old = next(n.get('Target') for n in original.xml(checker.relpath(v3_slides(original)[change['slide'] - 1])) if n.get('Type').endswith('/slideLayout'))
                next(n for n in p.xml(checker.relpath(part)) if n.get('Type').endswith('/slideLayout')).set('Target', old)
            add(f'layout-{index}-old-layout', old_layout)
            def wrong_mapping(p, change=change):
                part = v3_slides(p)[change['slide'] - 1]
                shape = reference_shape(p.xml(part), next(iter(change['placeholders'])))
                shape.find('p:nvSpPr/p:nvPr/p:ph', NS).set('idx', '999')
            add(f'layout-{index}-wrong-placeholder', wrong_mapping)
        for index, change in enumerate(spec.get('group_edits', [])):
            def flatten(p, change=change):
                root = p.xml(change['part'])
                group = reference_shape(root, change['group'])
                parent = next(n for n in root.iter() if group in list(n))
                at = list(parent).index(group)
                parent.remove(group)
                parent[at:at] = [n for n in group if n.tag not in (tag('p', 'nvGrpSpPr'), tag('p', 'grpSpPr'))]
            add(f'group-{index}-flattened', flatten)
            def transform(p, change=change):
                reference_shape(p.xml(change['part']), change['group']).find('p:grpSpPr/a:xfrm/a:off', NS).set('x', '99999')
            add(f'group-{index}-transform-changed', transform)
        if spec.get('table_structure'):
            change = spec['table_structure'][0]
            def bad_clone(p, change=change):
                table = p.xml(change['part']).findall('.//a:tbl', NS)[change['table'] - 1]
                rows = table.findall('a:tr', NS)
                rows[min(len(rows) - 1, change['index'] - 1)].set('h', '987654')
            add('table-wrong-template-format', bad_clone)
            def bad_grid(p, change=change):
                table = p.xml(change['part']).findall('.//a:tbl', NS)[change['table'] - 1]
                table.find('a:tblGrid', NS).remove(table.find('a:tblGrid', NS)[-1])
            add('table-wrong-column-grid', bad_grid)
            def undo_table(p, change=change):
                root = p.xml(change['part'])
                table = root.findall('.//a:tbl', NS)[change['table'] - 1]
                parent = next(n for n in root.iter() if table in list(n))
                parent[list(parent).index(table)] = copy.deepcopy(original.xml(change['part']).findall('.//a:tbl', NS)[change['table'] - 1])
            add('table-requested-structure-not-applied', undo_table)
        for index, change in enumerate(spec.get('duplicate_slide', [])):
            def shared_notes(p, change=change):
                parts = v3_slides(p)
                duplicate = parts[change['after']]
                original_notes = complex_related(original, v3_slides(original)[change['slide'] - 1], 'notesSlide')
                assert original_notes, 'duplicate matrix needs source notes'
                import posixpath
                next(n for n in p.xml(checker.relpath(duplicate)) if n.get('Type').endswith('/notesSlide')).set('Target', posixpath.relpath(original_notes, posixpath.dirname(duplicate)))
            if complex_related(original, v3_slides(original)[change['slide'] - 1], 'notesSlide'):
                add(f'duplicate-{index}-shared-notes', shared_notes)
            def dropped_link(p, change=change):
                duplicate = v3_slides(p)[change['after']]
                links = p.xml(checker.relpath(duplicate))
                link = next(n for n in links if n.get('Type').endswith('/slideLayout'))
                links.remove(link)
            add(f'duplicate-{index}-missing-layout', dropped_link)
            if any(n.get('Type').endswith('/image') for n in original.xml(checker.relpath(v3_slides(original)[change['slide'] - 1]))):
                def dropped_image(p, change=change):
                    duplicate = v3_slides(p)[change['after']]
                    links = p.xml(checker.relpath(duplicate))
                    links.remove(next(n for n in links if n.get('Type').endswith('/image')))
                add(f'duplicate-{index}-missing-image', dropped_image)
        for index, change in enumerate(spec.get('comments', [])):
            def no_parent(p, change=change):
                part = complex_related(p, v3_slides(p)[change['slide'] - 1], 'comments')
                node = next(n for n in p.xml(part) if n.findtext('p:text', namespaces=NS) == change['text'])
                node.remove(node.find('p:extLst', NS))
            add(f'comments-{index}-unthreaded', no_parent)
    return result


def complex_allowances(task, original):
    result = []
    if any(name.startswith('[trash]/') for name in original.parts):
        def drop_trash(p):
            for name in list(p.parts):
                if name.startswith('[trash]/'):
                    del p.parts[name]
                    p.xml_parts.pop(name, None)
            types = p.xml('[Content_Types].xml')
            types[:] = [n for n in types if not n.get('PartName', '').lstrip('/').startswith('[trash]/')]
        result.append(('ignorable-trash-removed', drop_trash))
    if task['format'] == 'xlsx' and task['expected'].get('chart'):
        def chart_ranges(p):
            for change in task['expected']['chart']:
                for series in p.xml(change['part']).findall('.//c:ser', NS):
                    for field in ('cat', 'val', 'xVal', 'yVal'):
                        target = series.find('c:' + field, NS)
                        if target is None:
                            continue
                        node = target.find('.//c:f', NS)
                        sheet, cells = node.text.rsplit('!', 1)
                        sheet = sheet[1:-1].replace("''", "'") if sheet.startswith("'") else sheet
                        node.text = "'" + sheet.replace("'", "''") + "'!" + cells.replace('$', '')
        result.append(('relative-quoted-chart-ranges', chart_ranges))
    if task['format'] == 'docx':
        def point_comment(p):
            for change in task['expected'].get('comments', []):
                if change['action'] == 'add':
                    paragraph = next(n for n in p.xml(change['part']).iter(tag('w', 'p')) if v3_text(n, 'w') == change['paragraph'])
                    part = complex_related(p, DOCUMENT, 'comments')
                    identity = next(n.get(tag('w', 'id')) for n in p.xml(part) if reference_comment_body(n) == change['text']
                                    and n.get(tag('w', 'author')) == change['author'])
                    paragraph[:] = [n for n in paragraph if n.tag not in (tag('w', 'commentRangeStart'), tag('w', 'commentRangeEnd'))
                                    or n.get(tag('w', 'id')) != identity]
        if any(n['action'] == 'add' for n in task['expected'].get('comments', [])):
            result.append(('point-comment-anchor', point_comment))
        def implicit_continue(p):
            for change in task['expected'].get('merge_cells', []):
                for node in p.xml(change['part']).iter(tag('w', 'vMerge')):
                    if node.get(tag('w', 'val')) == 'continue':
                        del node.attrib[tag('w', 'val')]
        if task['expected'].get('merge_cells'):
            result.append(('implicit-vmerge-continue', implicit_continue))
        if task['id'].endswith('op-matrix'):
            def inherited_list(p):
                for change in task['expected'].get('list_items', []):
                    paragraph = next(n for n in p.xml(change['part']).iter(tag('w', 'p')) if v3_text(n, 'w') == change['text'])
                    props = paragraph.find('w:pPr', NS)
                    props.find('w:pStyle', NS).set(tag('w', 'val'), 'CheckerList')
                    props.remove(props.find('w:numPr', NS))
            result.append(('style-inherited-list-numbering', inherited_list))
    if task['format'] == 'xlsx' and task['expected'].get('conditional_formats'):
        def pool_reorder(p):
            pool = p.xml('xl/styles.xml').find('s:dxfs', NS)
            pool.insert(0, ET.Element(tag('s', 'dxf')))
            for part in list(p.parts):
                if part.startswith('xl/') and part.endswith('.xml'):
                    for node in p.xml(part).iter():
                        if node.get('dxfId') is not None:
                            node.set('dxfId', str(int(node.get('dxfId')) + 1))
            pool.set('count', str(len(pool)))
        result.append(('differential-pool-reordered', pool_reorder))
        def split_range(p):
            change = task['expected']['conditional_formats'][0]
            c1, r1, c2, r2 = checker.range_coordinates(change['sqref'])
            assert r1 < r2
            root = p.xml(v3_sheet_map(p)[change['sheet']])
            group = [n for n in root.findall('s:conditionalFormatting', NS) if n.get('sqref') == change['sqref']][-1]
            other = copy.deepcopy(group)
            midpoint = (r1 + r2) // 2
            group.set('sqref', f'{checker.column_letters(c1)}{r1}:{checker.column_letters(c2)}{midpoint}')
            other.set('sqref', f'{checker.column_letters(c1)}{midpoint + 1}:{checker.column_letters(c2)}{r2}')
            priority = max(int(n.get('priority')) for n in root.findall('s:conditionalFormatting/s:cfRule', NS))
            for node in other:
                priority += 1
                node.set('priority', str(priority))
                for formula in node.findall('s:formula', NS):
                    formula.text = checker.translate_formula(formula.text, f'{checker.column_letters(c1)}{r1}',
                                                             f'{checker.column_letters(c1)}{midpoint + 1}')
            root.insert(list(root).index(group) + 1, other)
        first = task['expected']['conditional_formats'][0]['sqref']
        if len(task['expected']['conditional_formats']) == 1 and ':' in first and ' ' not in first and checker.range_coordinates(first)[1] < checker.range_coordinates(first)[3]:
            result.append(('split-conditional-range', split_range))
    return result


def validate_complex(directory):
    import json
    counts = {fmt: {key: [0, 0] for key in ('accepted', 'untouched_rejected', 'mutants_rejected')}
              for fmt in ('docx', 'xlsx', 'pptx')}
    failures = []
    def score(task, name, path, accept, kind, source):
        try:
            with patch.object(checker, 'fixture', return_value=source):
                checker.check(task, path)
            passed, reason = accept, 'accepted'
        except Exception as error:
            passed, reason = not accept, str(error)
        counts[task['format']][kind][0] += int(passed)
        counts[task['format']][kind][1] += 1
        if not passed:
            failures.append(f'{task["id"]} {name}: {reason}')
        print(f'{"OK" if passed else "FAIL"} complex {task["id"]} {name}: {reason if not passed else ""}', flush=True)
    def prepared(task):
        if not task.get('prepare'):
            return checker.fixture(task)
        path = directory / (task['id'] + '.input.' + task['format'])
        path.write_bytes(checker.prepare_input(task, checker.fixture(task).read_bytes()))
        return path
    cases = [(task, prepared(task)) for task in checker.tasks('v3', tier='complex')]
    cases += complex_cases(directory)
    for task, source in cases:
        if 'answer' in task['expected']:
            answer_directory = directory / task['id']
            answer_directory.mkdir(exist_ok=True)
            answer = answer_directory / 'answer.json'
            answer.unlink(missing_ok=True)
            score(task, 'untouched', answer, False, 'untouched_rejected', source)
            answer.write_text(json.dumps(task['expected']['answer']))
            score(task, 'reference', answer, True, 'accepted', source)
            for payload in [json.dumps({'unexpected_answer': task['expected']['answer']}),
                            json.dumps([task['expected']['answer']]), 'NaN', '{invalid']:
                answer.write_text(payload)
                score(task, 'corrupt-answer', answer, False, 'mutants_rejected', source)
            continue
        try:
            original, correct = Package(source), Package(source)
            complex_reference(correct, task)
            output = directory / (task['id'] + '.' + task['format'])
            complex_save(correct, output)
        except Exception as error:
            failures.append(f'{task["id"]} reference construction: {error}')
            counts[task['format']]['accepted'][1] += 1
            continue
        score(task, 'reference', output, True, 'accepted', source)
        score(task, 'untouched', source, False, 'untouched_rejected', source)
        try:
            partial = Package(source)
            partial_task = copy.deepcopy(task)
            del partial_task['expected'][next(reversed(partial_task['expected']))]
            if partial_task['expected']:
                complex_reference(partial, partial_task)
            path = directory / ('partial-' + task['id'] + '.' + task['format'])
            complex_save(partial, path)
            score(task, 'partial-edit', path, False, 'mutants_rejected', source)
        except Exception as error:
            failures.append(f'{task["id"]} partial construction: {error}')
            counts[task['format']]['mutants_rejected'][1] += 1
        mutations = v3_mutants(correct, task, original) + complex_mutants(task, original)
        for name, mutate in mutations + complex_allowances(task, original):
            accept = (name, mutate) not in mutations
            kind = 'accepted' if accept else 'mutants_rejected'
            try:
                package = Package(output)
                mutate(package)
                path = directory / ('complex-variant-' + task['id'] + '.' + task['format'])
                complex_save(package, path)
                score(task, name, path, accept, kind, source)
            except Exception as error:
                failures.append(f'{task["id"]} {name} construction: {error}')
                counts[task['format']][kind][1] += 1
    for fmt in counts:
        base = next(task for task in checker.tasks('v3', tier='complex') if task['format'] == fmt)
        answer_task = dict(base, id=fmt + '-v3-complex-answer', expected={'answer': {'orientation': 'landscape', 'pages': [1, 3]}})
        answer_directory = directory / (fmt + '-answers')
        answer_directory.mkdir(exist_ok=True)
        path = answer_directory / 'answer.json'
        for name, payload, accept, kind in [
                ('missing', None, False, 'untouched_rejected'),
                ('reference', json.dumps(answer_task['expected']['answer']), True, 'accepted'),
                ('wrong-value', '{"orientation":"portrait","pages":[1,3]}', False, 'mutants_rejected'),
                ('extra-key', '{"orientation":"landscape","pages":[1,3],"extra":1}', False, 'mutants_rejected'),
                ('boolean-number', '{"orientation":"landscape","pages":[true,3]}', False, 'mutants_rejected'),
                ('duplicate-key', '{"orientation":"landscape","orientation":"landscape","pages":[1,3]}', False, 'mutants_rejected'),
                ('invalid-json', '{broken', False, 'mutants_rejected')]:
            path.unlink(missing_ok=True)
            if payload is not None:
                path.write_text(payload)
            score(answer_task, name, path, accept, kind, checker.fixture(base))
    for fmt, values in counts.items():
        print('V3 COMPLEX ' + fmt.upper() + ' ' + ' '.join(f'{key}={p}/{n}' for key, (p, n) in values.items()))
    totals = {key: [sum(values[key][i] for values in counts.values()) for i in (0, 1)]
              for key in ('accepted', 'untouched_rejected', 'mutants_rejected')}
    print('V3 COMPLEX VALIDATE ' + ' '.join(f'{key}={p}/{n}' for key, (p, n) in totals.items()), flush=True)
    assert not failures, '\n'.join(failures)


def validate_v3(directory):
    counts = {fmt: {'accepted':[0,0], 'untouched_rejected':[0,0], 'mutants_rejected':[0,0]} for fmt in ('docx','xlsx','pptx')}
    failures = []
    def score(task, name, path, accept, kind):
        try:
            checker.check(task, path)
            passed = accept
            reason = 'accepted'
        except Exception as error:
            passed = not accept
            reason = str(error)
        count = counts[task['format']][kind]
        count[0] += int(passed)
        count[1] += 1
        if not passed:
            failures.append(f'{task["id"]} {name}: {reason}')
        print(f'{"OK" if passed else "FAIL"} {task["id"]} {name}: {reason if not passed else ""}', flush=True)
    checked_allowances = set()
    for task in checker.tasks('v3'):
        source = checker.fixture(task)
        original = Package(source)
        correct = Package(source)
        try:
            v3_correct(correct, task)
        except Exception as error:
            failures.append(f'{task["id"]} reference construction: {error}')
            counts[task['format']]['accepted'][1] += 1
            print(f'FAIL {failures[-1]}', flush=True)
            continue
        output = directory / (task['id'] + '.' + task['format'])
        v3_save(correct, output)
        score(task, 'reference', output, True, 'accepted')
        score(task, 'untouched', source, False, 'untouched_rejected')
        try:
            partial = v3_partial(source, task, output)
            path = directory / ('partial-' + task['id'] + '.' + task['format'])
            v3_save(partial, path)
            score(task, 'partial-edit', path, False, 'mutants_rejected')
        except Exception as error:
            counts[task['format']]['mutants_rejected'][1] += 1
            failures.append(f'{task["id"]} partial construction: {error}')
            print('FAIL ' + failures[-1], flush=True)
        for name, mutate in v3_mutants(correct, task, original):
            package = Package(output)
            try:
                mutate(package)
                path = directory / ('mutant-' + task['id'] + '.' + task['format'])
                v3_save(package, path)
                score(task, name, path, False, 'mutants_rejected')
            except Exception as error:
                counts[task['format']]['mutants_rejected'][1] += 1
                failures.append(f'{task["id"]} mutant construction {name}: {error}')
                print('FAIL ' + failures[-1], flush=True)
        for name, mutate, accept in [*v3_calculation_variants(task), *v3_fidelity_variants(task, original)]:
            package = Package(output)
            kind = 'accepted' if accept else 'mutants_rejected'
            try:
                mutate(package)
                path = directory / (name + '-' + task['id'] + '.' + task['format'])
                v3_save(package, path)
                score(task, name, path, accept, kind)
            except Exception as error:
                counts[task['format']][kind][1] += 1
                failures.append(f'{task["id"]} variant construction {name}: {error}')
                print('FAIL ' + failures[-1], flush=True)
        allowances = v3_allowances(task)
        if task['format'] in checked_allowances:
            allowances = [(name, mutate) for name, mutate in allowances if name == 'updated-slide-title-metadata']
        checked_allowances.add(task['format'])
        for name, mutate in allowances:
            package = Package(output)
            try:
                mutate(package)
                path = directory / ('allowance-' + task['id'] + '.' + task['format'])
                v3_save(package, path)
                score(task, name, path, True, 'accepted')
            except Exception as error:
                counts[task['format']]['accepted'][1] += 1
                failures.append(f'{task["id"]} allowance construction {name}: {error}')
                print('FAIL ' + failures[-1], flush=True)
    for fmt, values in counts.items():
        print('V3 ' + fmt.upper() + ' ' + ' '.join(f'{key}={p}/{n}' for key,(p,n) in values.items()))
    totals = {key: [sum(values[key][i] for values in counts.values()) for i in (0,1)] for key in ('accepted','untouched_rejected','mutants_rejected')}
    print('V3 VALIDATE ' + ' '.join(f'{key}={p}/{n}' for key,(p,n) in totals.items()), flush=True)
    print('V3 BASIC VALIDATE ' + ' '.join(f'{key}={p}/{n}' for key,(p,n) in totals.items()), flush=True)
    assert not failures, '\n'.join(failures)
    validate_complex(directory)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--output-dir', type=Path)
    parser.add_argument('--suite', choices=('v2','v3'), default='v2')
    args = parser.parse_args()
    validate_suite = validate_v3 if args.suite == 'v3' else validate
    if args.output_dir:
        args.output_dir.mkdir(parents=True, exist_ok=True)
        validate_suite(args.output_dir)
    else:
        with tempfile.TemporaryDirectory(prefix='office-checkers-') as temporary:
            validate_suite(Path(temporary))


if __name__ == '__main__':
    main()
