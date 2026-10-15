import argparse
import contextlib
import copy
import importlib.util
import tempfile
import xml.etree.ElementTree as ET
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
        spec['insert_paragraphs']['texts'] = spec['insert_paragraphs']['texts'][:1]
    elif 'format_phrase' in spec:
        change = spec['format_phrase']
        change['phrase'] = change['phrase'][:max(1, len(change['phrase']) // 2)]
    else:
        spec.clear()
    partial = Package(source)
    v3_correct(partial, partial_task)
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
    assert not failures, '\n'.join(failures)


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
