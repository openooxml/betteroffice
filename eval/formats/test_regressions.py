import importlib.util
import json
import os
import signal
import subprocess
import sys
import tempfile
import time
import unittest
from io import BytesIO
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from zipfile import ZipFile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location('regression_runner', ROOT / 'run.py')
runner = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(runner)


class WorkspaceTests(unittest.TestCase):
    def test_agent_paths_cannot_redirect_harness_reads_or_writes(self):
        for target in ('answer.json', 'input.xlsx', 'output.xlsx'):
            for kind in ('symlink', 'directory', 'fifo', 'parent-symlink'):
                with self.subTest(target=target, kind=kind), tempfile.TemporaryDirectory() as temporary:
                    root = Path(temporary)
                    outside = root / 'outside'
                    outside.mkdir()
                    victim = outside / Path(target).name
                    victim.write_bytes(b'host content')
                    fixture = root / 'fixture.xlsx'
                    fixture.write_bytes(b'fixture')
                    name = 'nested/' + target if kind == 'parent-symlink' else target
                    task = {'id': 'safe-path', 'format': 'xlsx', 'input': name if target == 'input.xlsx' else 'input.xlsx',
                            'output': name if target == 'output.xlsx' else 'output.xlsx', 'prompt': 'test'}
                    if target == 'answer.json':
                        task.update(tier='complex', expected={'answer': 4})
                    else:
                        task['expected'] = {}
                    def process(command, cwd, stdout, **kwargs):
                        workspace = Path(cwd)
                        path = workspace / name
                        if kind == 'parent-symlink':
                            path.unlink(missing_ok=True)
                            path.parent.rmdir() if path.parent.exists() else None
                            path.parent.symlink_to(outside, target_is_directory=True)
                        else:
                            path.unlink(missing_ok=True)
                            if kind == 'symlink':
                                path.symlink_to(victim)
                            elif kind == 'fifo':
                                os.mkfifo(path)
                            else:
                                path.mkdir()
                        if target == 'answer.json' and kind == 'parent-symlink':
                            (workspace / 'answer.json').symlink_to(path)
                        for item in ({'id': 'tool', 'type': 'command_execution', 'command': 'true'},
                                     {'id': 'final', 'type': 'agent_message', 'text': '4'}):
                            stdout.write(json.dumps({'type': 'item.completed', 'item': item}) + '\n')
                        return SimpleNamespace(returncode=0, stdin=None, stdout=None, poll=lambda: 0, wait=lambda timeout: 0)
                    with patch.object(runner.checker, 'fixture', return_value=fixture), \
                            patch.object(runner.checker, 'prepare_input', side_effect=lambda task, data: data), \
                            patch.object(runner.checker, 'check'), \
                            patch.object(runner, 'command_line', return_value=['fake']), \
                            patch.object(runner.subprocess, 'Popen', side_effect=process):
                        record = runner.run_task(task, None, 'model', root / 'results', arm='generic')
                    self.assertEqual(victim.read_bytes(), b'host content')
                    self.assertFalse(record['pass'])
                    self.assertEqual(record['failure_type'], 'unsafe-workspace-file')
                    self.assertFalse((root / 'results' / task['id'] / target).exists())


class ConditionalFormatTests(unittest.TestCase):
    def normalized(self, ranges):
        checker = runner.checker
        sheet = ET.Element(checker.q('s', 'worksheet'))
        for priority, (sqref, formula) in enumerate(ranges, 1):
            group = ET.SubElement(sheet, checker.q('s', 'conditionalFormatting'), sqref=sqref)
            rule = ET.SubElement(group, checker.q('s', 'cfRule'), type='expression', priority=str(priority))
            ET.SubElement(rule, checker.q('s', 'formula')).text = formula
        styles = ET.Element(checker.q('s', 'styleSheet'))
        package = SimpleNamespace(root=lambda name: styles if name == 'xl/styles.xml' else sheet,
                                  namelist=lambda: ['xl/styles.xml', 'xl/worksheets/sheet1.xml'])
        with patch.object(checker, 'related', return_value='xl/styles.xml'), \
                patch.object(checker, 'sheet_map', return_value={'Sheet': 'xl/worksheets/sheet1.xml'}):
            checker.normalize_complex_xlsx(package, {})
        return checker.canonical(sheet)

    def test_cx07_split_range_requires_shifted_relative_formula(self):
        task = next(t for t in runner.checker.tasks('v3', 'complex') if t['id'] == 'xlsx-v3-cx-07-green-recovery-sort')
        change = task['expected']['conditional_formats'][0]
        formula = change['formulas'][0]
        self.assertEqual(change['sqref'], 'E4:E30')
        expected = self.normalized([('E4:E30', formula)])
        self.assertNotEqual(expected, self.normalized([('E4:E17', formula), ('E18:E30', formula)]))
        shifted = runner.checker.translate_formula(formula, 'E4', 'E18')
        self.assertEqual(expected, self.normalized([('E4:E17', formula), ('E18:E30', shifted)]))

    def test_mixed_absolute_references_and_quoted_text(self):
        formula = 'AND(E4=$A$1,$B4=E$2,C3="E4")'
        expected = self.normalized([('E4:F30', formula)])
        shifted = runner.checker.translate_formula(formula, 'E4', 'F18')
        self.assertEqual(expected, self.normalized([('E4:F17', formula), ('E18:E30', runner.checker.translate_formula(formula, 'E4', 'E18')), ('F18:F30', shifted)]))
        self.assertNotEqual(expected, self.normalized([('E4:F17', formula), ('E18:F30', formula)]))

    def test_reference_before_range_origin_is_valid(self):
        self.assertEqual(self.normalized([('E4:E30', 'A1>0')]),
                         self.normalized([('E4:E17', 'A1>0'), ('E18:E30', 'A15>0')]))


RELS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"')
P_NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"')
THREAD_URI = '{C676402C-5697-4E1C-873F-D02D1690AC5C}'


def archive(parts):
    buffer = BytesIO()
    with ZipFile(buffer, 'w') as output:
        for name, data in parts.items():
            output.writestr(name, data)
    buffer.seek(0)
    return ZipFile(buffer)


def content_types(overrides):
    items = ''.join(f'<Override PartName="/{name}" ContentType="{kind}"/>' for name, kind in overrides.items())
    return ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            f'<Default Extension="xml" ContentType="application/xml"/>{items}</Types>')


def relationships(links):
    items = ''.join(f'<Relationship Id="rId{index}" Type="{kind}" Target="{target}"/>'
                    for index, (kind, target) in enumerate(links, 1))
    return f'<Relationships xmlns="{PACKAGE_RELS}">{items}</Relationships>'


def word_package(body, comments, extended):
    main = 'application/vnd.openxmlformats-officedocument.wordprocessingml'
    return archive({
        '[Content_Types].xml': content_types({'word/document.xml': f'{main}.document.main+xml',
                                              'word/styles.xml': f'{main}.styles+xml',
                                              'word/comments.xml': f'{main}.comments+xml',
                                              'word/commentsExtended.xml': f'{main}.commentsExtended+xml'}),
        '_rels/.rels': relationships([(f'{RELS}/officeDocument', 'word/document.xml')]),
        'word/_rels/document.xml.rels': relationships([
            (f'{RELS}/styles', 'styles.xml'), (f'{RELS}/comments', 'comments.xml'),
            ('http://schemas.microsoft.com/office/2011/relationships/commentsExtended', 'commentsExtended.xml')]),
        'word/document.xml': f'<w:document {W_NS}><w:body>{body}<w:sectPr/></w:body></w:document>',
        'word/styles.xml': f'<w:styles {W_NS}/>',
        'word/comments.xml': f'<w:comments {W_NS}>{comments}</w:comments>',
        'word/commentsExtended.xml': f'<w15:commentsEx {W_NS}>{extended}</w15:commentsEx>',
    })


def word_comment(identity, author, text, paragraph):
    return (f'<w:comment w:id="{identity}" w:author="{author}"><w:p w14:paraId="{paragraph}">'
            f'<w:r><w:t>{text}</w:t></w:r></w:p></w:comment>')


def text_run(text):
    space = ' xml:space="preserve"' if text != text.strip() else ''
    return f'<w:r><w:t{space}>{text}</w:t></w:r>'


def anchored(identity, text):
    return (f'<w:commentRangeStart w:id="{identity}"/>{text_run(text)}<w:commentRangeEnd w:id="{identity}"/>'
            f'<w:r><w:commentReference w:id="{identity}"/></w:r>')


def sheet_package(rows, auto_filter=None, names=''):
    main = 'application/vnd.openxmlformats-officedocument.spreadsheetml'
    cells = ''.join(f'<row r="{r}">' + ''.join(f'<c r="{c}{r}"><v>{v}</v></c>' for c, v in zip('AB', values)) + '</row>'
                    for r, values in enumerate(rows, 1))
    filtered = f'<autoFilter ref="{auto_filter}"/>' if auto_filter else ''
    defined = f'<definedNames>{names}</definedNames>' if names else ''
    return archive({
        '[Content_Types].xml': content_types({'xl/workbook.xml': f'{main}.sheet.main+xml',
                                              'xl/worksheets/sheet1.xml': f'{main}.worksheet+xml',
                                              'xl/styles.xml': f'{main}.styles+xml'}),
        '_rels/.rels': relationships([(f'{RELS}/officeDocument', 'xl/workbook.xml')]),
        'xl/_rels/workbook.xml.rels': relationships([(f'{RELS}/worksheet', 'worksheets/sheet1.xml'), (f'{RELS}/styles', 'styles.xml')]),
        'xl/workbook.xml': (f'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{RELS}">'
                            f'<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>{defined}</workbook>'),
        'xl/worksheets/sheet1.xml': (f'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                                     f'<sheetData>{cells}</sheetData>{filtered}</worksheet>'),
        'xl/styles.xml': ('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                          '<fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders>'
                          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                          '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>'),
    })


def slide_package(authors, comments):
    main = 'application/vnd.openxmlformats-officedocument.presentationml'
    return archive({
        '[Content_Types].xml': content_types({'ppt/presentation.xml': f'{main}.presentation.main+xml',
                                              'ppt/slides/slide1.xml': f'{main}.slide+xml',
                                              'ppt/commentAuthors.xml': f'{main}.commentAuthors+xml',
                                              'ppt/comments/comment1.xml': f'{main}.comments+xml'}),
        '_rels/.rels': relationships([(f'{RELS}/officeDocument', 'ppt/presentation.xml')]),
        'ppt/_rels/presentation.xml.rels': relationships([(f'{RELS}/slide', 'slides/slide1.xml'),
                                                          (f'{RELS}/commentAuthors', 'commentAuthors.xml')]),
        'ppt/slides/_rels/slide1.xml.rels': relationships([(f'{RELS}/comments', '../comments/comment1.xml')]),
        'ppt/presentation.xml': f'<p:presentation {P_NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>',
        'ppt/slides/slide1.xml': f'<p:sld {P_NS}><p:cSld><p:spTree/></p:cSld></p:sld>',
        'ppt/commentAuthors.xml': f'<p:cmAuthorLst {P_NS}>{authors}</p:cmAuthorLst>',
        'ppt/comments/comment1.xml': f'<p:cmLst {P_NS}>{comments}</p:cmLst>',
    })


def slide_comment(author, index, text, thread=''):
    extension = f'<p:extLst><p:ext uri="{THREAD_URI}">{thread}</p:ext></p:extLst>' if thread else ''
    return (f'<p:cm authorId="{author}" dt="2020-01-01T00:00:00.000" idx="{index}"><p:pos x="10" y="10"/>'
            f'<p:text>{text}</p:text>{extension}</p:cm>')


def layout_shape(identity, name, placeholder, text=''):
    body = f'<p:txBody><a:bodyPr/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody>' if text else '<p:txBody><a:bodyPr/><a:p/></p:txBody>'
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{identity}" name="{name}"/><p:cNvSpPr/><p:nvPr>{placeholder}</p:nvPr></p:nvSpPr>'
            f'<p:spPr/>{body}</p:sp>')


def layout_slide_package(slide_placeholder, layout='slideLayout1.xml'):
    main = 'application/vnd.openxmlformats-officedocument.presentationml'
    layouts = {
        'slideLayout1.xml': ('Title Only', layout_shape(2, 'Title 1', '<p:ph type="title"/>')),
        'slideLayout2.xml': ('Title and Content', layout_shape(2, 'Title 1', '<p:ph type="title"/>')
                             + layout_shape(3, 'Content Placeholder 2', '<p:ph idx="1" hasCustomPrompt="1"/>', 'Add notes')),
    }
    parts = {
        '[Content_Types].xml': content_types({'ppt/presentation.xml': f'{main}.presentation.main+xml',
                                              'ppt/slides/slide1.xml': f'{main}.slide+xml',
                                              **{f'ppt/slideLayouts/{name}': f'{main}.slideLayout+xml' for name in layouts}}),
        '_rels/.rels': relationships([(f'{RELS}/officeDocument', 'ppt/presentation.xml')]),
        'ppt/_rels/presentation.xml.rels': relationships([(f'{RELS}/slide', 'slides/slide1.xml')]),
        'ppt/presentation.xml': f'<p:presentation {P_NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>',
        'ppt/slides/_rels/slide1.xml.rels': relationships([(f'{RELS}/slideLayout', f'../slideLayouts/{layout}')]),
        'ppt/slides/slide1.xml': (f'<p:sld {P_NS}><p:cSld><p:spTree>'
                                  + layout_shape(2, 'Title 1', '<p:ph type="title"/>', 'Heading')
                                  + layout_shape(3, 'Subtitle 2', slide_placeholder, 'Review copy')
                                  + '</p:spTree></p:cSld></p:sld>'),
    }
    for name, (title, shapes) in layouts.items():
        parts[f'ppt/slideLayouts/{name}'] = f'<p:sldLayout {P_NS}><p:cSld name="{title}"><p:spTree>{shapes}</p:spTree></p:cSld></p:sldLayout>'
    return archive(parts)


def table_frame(identity, columns, rows, size, offset=(100, 200)):
    grid = ''.join(f'<a:gridCol w="{width}"/>' for width, _ in columns)
    cells = ''.join(f'<a:tr h="{height}">' + ''.join(f'<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>{row}{column}</a:t></a:r></a:p></a:txBody></a:tc>'
                                                    for _, column in columns) + '</a:tr>'
                    for height, row in rows)
    return (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{identity}" name="Table {identity}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
            f'<p:xfrm><a:off x="{offset[0]}" y="{offset[1]}"/><a:ext cx="{size[0]}" cy="{size[1]}"/></p:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            f'<a:tbl><a:tblGrid>{grid}</a:tblGrid>{cells}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>')


def table_slide_package(frames):
    main = 'application/vnd.openxmlformats-officedocument.presentationml'
    return archive({
        '[Content_Types].xml': content_types({'ppt/presentation.xml': f'{main}.presentation.main+xml',
                                              'ppt/slides/slide1.xml': f'{main}.slide+xml'}),
        '_rels/.rels': relationships([(f'{RELS}/officeDocument', 'ppt/presentation.xml')]),
        'ppt/_rels/presentation.xml.rels': relationships([(f'{RELS}/slide', 'slides/slide1.xml')]),
        'ppt/presentation.xml': f'<p:presentation {P_NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>',
        'ppt/slides/slide1.xml': f'<p:sld {P_NS}><p:cSld><p:spTree>{frames}</p:spTree></p:cSld></p:sld>',
    })


def chart_workbook(chart, series='Old'):
    package = sheet_package([(1, 2), (3, 4)])
    parts = {name: package.read(name) for name in package.namelist()}
    main = 'application/vnd.openxmlformats-officedocument'
    parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace(
        b'</Types>', f'<Override PartName="/xl/charts/chart1.xml" ContentType="{main}.drawingml.chart+xml"/></Types>'.encode())
    parts['xl/charts/chart1.xml'] = ('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
                                     'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                                     f'<c:chart>{chart}<c:plotArea><c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/>'
                                     f'<c:tx><c:v>{series}</c:v></c:tx><c:cat><c:strRef><c:f>Data!$A$1:$A$2</c:f></c:strRef></c:cat>'
                                     '<c:val><c:numRef><c:f>Data!$B$1:$B$2</c:f></c:numRef></c:val></c:ser></c:lineChart></c:plotArea>'
                                     '</c:chart></c:chartSpace>')
    return archive(parts)


def chart_title(extra=''):
    return ('<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Totals</a:t></a:r></a:p></c:rich></c:tx>'
            f'{extra}</c:title>')


def complex_task(fmt, expected):
    return {'id': 'synthetic', 'format': fmt, 'tier': 'complex', 'suite': 'v3', 'expected': expected}


class DependentReferenceCacheTests(unittest.TestCase):
    def package(self, edited=False, caches=None, full_calc=False, formula="'FY27 D'!C6", mutate=None,
                values=('GAP Coordinator', 'Environmental Technician')):
        checker = runner.checker
        with sheet_package([]) as base:
            parts = {name: base.read(name) for name in base.namelist()}
        root = ET.fromstring(parts['xl/workbook.xml'])
        sheets = root.find('s:sheets', checker.NS)
        sheets[0].set('name', 'FY27 D')
        ET.SubElement(sheets, checker.q('s', 'sheet'), {'name': 'FY27 Exp', 'sheetId': '2', f'{{{RELS}}}id': 'rId3'})
        calculation = ET.SubElement(root, checker.q('s', 'calcPr'), calcId='191029')
        if full_calc:
            calculation.set('fullCalcOnLoad', '1')
        parts['xl/workbook.xml'] = ET.tostring(root)
        root = ET.fromstring(parts['xl/_rels/workbook.xml.rels'])
        ET.SubElement(root, f'{{{PACKAGE_RELS}}}Relationship', Id='rId3', Type=f'{RELS}/worksheet', Target='worksheets/sheet2.xml')
        parts['xl/_rels/workbook.xml.rels'] = ET.tostring(root)
        root = ET.fromstring(parts['[Content_Types].xml'])
        ET.SubElement(root, '{http://schemas.openxmlformats.org/package/2006/content-types}Override',
                      PartName='/xl/worksheets/sheet2.xml',
                      ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
        parts['[Content_Types].xml'] = ET.tostring(root)
        def cell(data, address, value, expression=None, properties=None):
            row = ET.SubElement(data, checker.q('s', 'row'), r=''.join(filter(str.isdigit, address)))
            kind = 'b' if isinstance(value, bool) else 'n' if isinstance(value, (int, float)) else 'str'
            node = ET.SubElement(row, checker.q('s', 'c'), r=address, t=kind)
            if expression is not None:
                ET.SubElement(node, checker.q('s', 'f'), properties or {}).text = expression
            ET.SubElement(node, checker.q('s', 'v')).text = str(int(value)) if isinstance(value, bool) else str(value)
        inputs = ET.Element(checker.q('s', 'worksheet'))
        data = ET.SubElement(inputs, checker.q('s', 'sheetData'))
        for address, value in (('C6', values[0]), ('C7', 'Coordinate the approved GAP work plan.'),
                               ('C8', values[1]), ('C9', 'Support solid waste monitoring and reporting.')):
            old = 'Staff Title' if address in ('C6', 'C8') else 'Description of duties'
            if isinstance(value, (int, float, bool)):
                old = False if isinstance(value, bool) else 1
            cell(data, address, value if edited else old)
        cell(data, 'C10', 'Unedited')
        outputs = ET.Element(checker.q('s', 'worksheet'))
        data = ET.SubElement(outputs, checker.q('s', 'sheetData'))
        old = [False if isinstance(value, bool) else 1 if isinstance(value, (int, float)) else 'Staff Title' for value in values]
        for address, value, expression, properties in (
                ('B7', old[0], formula, {}), ('B11', old[1], "'FY27 D'!$C$8", {}),
                ('B12', 2, '1+1', {}), ('B13', 'Unedited', "'FY27 D'!C10", {}),
                ('B14', old[0], "'FY27 D'!C6", {'t': 'array', 'ref': 'B14'}),
                ('B15', old[0], "'FY27 D'!C6", {'t': 'shared', 'si': '0', 'ref': 'B15'})):
            cell(data, address, (caches or {}).get(address, value), expression, properties)
        if mutate:
            mutate(inputs, outputs)
        parts['xl/worksheets/sheet1.xml'] = ET.tostring(inputs)
        parts['xl/worksheets/sheet2.xml'] = ET.tostring(outputs)
        return archive(parts)

    def task(self, values=('GAP Coordinator', 'Environmental Technician')):
        return {'format': 'xlsx', 'expected': {'cells': {'FY27 D': {
            'C6': {'value': values[0]}, 'C7': {'value': 'Coordinate the approved GAP work plan.'},
            'C8': {'value': values[1]}, 'C9': {'value': 'Support solid waste monitoring and reporting.'}}}}}

    def verdict(self, accepted, **kwargs):
        with self.package(formula=kwargs.get('formula', "'FY27 D'!C6"), values=kwargs.get('values', ('GAP Coordinator', 'Environmental Technician'))) as source, \
                self.package(edited=True, **kwargs) as output:
            if accepted:
                runner.checker.check_real(self.task(kwargs.get('values', ('GAP Coordinator', 'Environmental Technician'))), source, output)
            else:
                with self.assertRaises(ValueError):
                    runner.checker.check_real(self.task(kwargs.get('values', ('GAP Coordinator', 'Environmental Technician'))), source, output)

    def test_exact_recomputation_and_legacy_caches(self):
        for full_calc in (False, True):
            for caches in (None, {'B7': 'GAP Coordinator', 'B11': 'Environmental Technician'}):
                with self.subTest(full_calc=full_calc, caches=caches):
                    self.verdict(True, full_calc=full_calc, caches=caches)

    def test_sheet_names_match_case_insensitively_like_excel(self):
        for formula in ("'fy27 d'!C6", "'FY27 d'!$C$6"):
            with self.subTest(formula=formula):
                self.verdict(True, formula=formula, caches={'B7': 'GAP Coordinator'})
                self.verdict(False, formula=formula, caches={'B7': 'Wrong'})

    def test_invalid_caches_stay_strict(self):
        for caches in ({'B7': 'Wrong'}, {'B7': 'Environmental Technician', 'B11': 'GAP Coordinator'},
                       {'B7': ''}, {'B12': 3}, {'B13': 'Changed'}, {'B14': 'GAP Coordinator'},
                       {'B15': 'GAP Coordinator'}):
            for full_calc in (False, True):
                with self.subTest(caches=caches, full_calc=full_calc):
                    self.verdict(False, caches=caches, full_calc=full_calc)

    def test_changed_formula_properties_type_attributes_and_input_fail(self):
        ns = runner.checker.NS
        mutations = (
            lambda inputs, outputs: outputs.find('.//s:c[@r="B7"]/s:f', ns).__setattr__('text', "'FY27 D'!C8"),
            lambda inputs, outputs: outputs.find('.//s:c[@r="B7"]/s:f', ns).set('aca', '1'),
            lambda inputs, outputs: outputs.find('.//s:c[@r="B7"]', ns).set('t', 'e'),
            lambda inputs, outputs: outputs.find('.//s:c[@r="B7"]', ns).set('cm', '1'),
            lambda inputs, outputs: inputs.find('.//s:c[@r="C6"]/s:v', ns).__setattr__('text', 'Wrong'),
        )
        for index, mutate in enumerate(mutations):
            with self.subTest(mutation=index):
                self.verdict(False, caches={'B7': 'GAP Coordinator'}, mutate=mutate)

    def test_functions_ranges_operators_and_external_references_stay_strict(self):
        for formula in ("SUM('FY27 D'!C6:C7)", "UPPER('FY27 D'!C6)", "'FY27 D'!C6:C7",
                        "'FY27 D'!C6&\"\"", "'[other.xlsx]FY27 D'!C6"):
            with self.subTest(formula=formula):
                self.verdict(False, formula=formula, caches={'B7': 'GAP Coordinator'})

    def test_number_and_bool_recomputation_is_typed(self):
        for values, caches, accepted in (
                ((2.5, 4), {'B7': 2.500000000001, 'B11': 4}, True),
                ((2.5, 4), {'B7': 2.6}, False),
                ((True, True), {'B7': True, 'B11': True}, True),
                ((True, True), {'B7': 1}, False)):
            with self.subTest(values=values, caches=caches):
                self.verdict(accepted, values=values, caches=caches)

    def test_single_reference_parser(self):
        parse = runner.checker.single_cell_reference
        for formula in ('C6', '=$C$6', 'FY27!C6', "'FY27 D'!$C6", "'O''Brien'!C6"):
            with self.subTest(formula=formula):
                self.assertIsNotNone(parse(formula, 'FY27 D'))
        self.assertEqual(parse("'O''Brien'!$C$6", 'Sheet'), ("O'Brien", 'C6'))
        for formula in ('C6:C7', 'UPPER(C6)', '[other.xlsx]Sheet!C6', "'[other.xlsx]Sheet'!C6", 'XFE1', 'A1048577'):
            with self.subTest(formula=formula):
                self.assertIsNone(parse(formula, 'Sheet'))


class RenamedSheetTitleTests(unittest.TestCase):
    def package(self, sheet, titles, names=None):
        names = names or (f"{sheet_reference(sheet)}!$A$1", f"{sheet_reference(sheet)}!$A$2")
        with sheet_package([(1, 2)], names=f'<definedName name="Version">{names[0]}</definedName>'
                           f'<definedName name="Notes" localSheetId="0">{names[1]}</definedName>') as base:
            parts = {name: base.read(name) for name in base.namelist()}
        parts['xl/workbook.xml'] = parts['xl/workbook.xml'].replace(b'name="Data"', f'name="{sheet}"'.encode())
        vt = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
        parts['docProps/app.xml'] = (f'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="{vt}">'
                                     '<HeadingPairs><vt:vector size="4" baseType="variant">'
                                     '<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant>'
                                     '<vt:variant><vt:lpstr>Named Ranges</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant>'
                                     '</vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="3" baseType="lpstr">'
                                     + ''.join(f'<vt:lpstr>{title}</vt:lpstr>' for title in titles) + '</vt:vector></TitlesOfParts></Properties>')
        return archive(parts)

    def test_rename_keeps_a_global_name_title_that_matches_the_old_sheet(self):
        task = {'format': 'xlsx', 'suite': 'v3', 'expected': {'rename': {'Version': 'Version History'}}}
        for name, titles, accepted in (
                ('global title kept, local title requalified', ('Version History', 'Version', "'Version History'!Notes"), True),
                ('global title renamed', ('Version History', 'Version History', "'Version History'!Notes"), False),
                ('worksheet title kept', ('Version', 'Version', "'Version History'!Notes"), False),
                ('local title kept', ('Version History', 'Version', 'Version!Notes'), False)):
            with self.subTest(output=name), self.package('Version', ('Version', 'Version', 'Version!Notes')) as source, \
                    self.package('Version History', titles) as output:
                if accepted:
                    runner.checker.check_real(task, source, output)
                else:
                    with self.assertRaises(ValueError):
                        runner.checker.check_real(task, source, output)

    def test_rename_to_a_plain_name_accepts_a_quoted_or_bare_local_title(self):
        task = {'format': 'xlsx', 'suite': 'v3', 'expected': {'rename': {'Version': 'Reg_01'}}}
        names = ("'Reg_01'!$A$1", "'Reg_01'!$A$2")
        for name, titles, accepted in (
                ('bare qualifier', ('Reg_01', 'Version', 'Reg_01!Notes'), True),
                ('quoted qualifier', ('Reg_01', 'Version', "'Reg_01'!Notes"), True),
                ('quoted wrong sheet', ('Reg_01', 'Version', "'Reg_02'!Notes"), False),
                ('bare old sheet', ('Reg_01', 'Version', 'Version!Notes'), False),
                ('quoted wrong name', ('Reg_01', 'Version', "'Reg_01'!Note"), False)):
            with self.subTest(output=name), self.package('Version', ('Version', 'Version', 'Version!Notes')) as source, \
                    self.package('Reg_01', titles, names) as output:
                if accepted:
                    runner.checker.check_real(task, source, output)
                else:
                    with self.assertRaises(ValueError):
                        runner.checker.check_real(task, source, output)


def sheet_reference(name):
    return name if name.isidentifier() else "'" + name.replace("'", "''") + "'"


class OfficeEquivalentTests(unittest.TestCase):
    def assert_verdicts(self, task, source, outputs):
        for name, (output, accepted) in outputs.items():
            with self.subTest(output=name):
                try:
                    runner.checker.check_complex(task, source(), output())
                    verdict = None
                except ValueError as error:
                    verdict = str(error)
                if accepted:
                    self.assertIsNone(verdict)
                else:
                    self.assertIsNotNone(verdict, 'checker accepted a non-equivalent output')

    def test_word_comment_reply_may_share_its_parent_anchors(self):
        parent = word_comment(0, 'Reviewer', 'Check this', '10000001')
        reply = word_comment(1, 'Editor', 'Checked', '10000002')
        threaded = '<w15:commentEx w15:paraId="10000001"/><w15:commentEx w15:paraId="10000002" w15:paraIdParent="10000001"/>'
        unthreaded = '<w15:commentEx w15:paraId="10000001"/><w15:commentEx w15:paraId="10000002"/>'
        def body(*anchors):
            return f'<w:p>{text_run("Alpha ")}{"".join(anchors)}{text_run(" gamma.")}</w:p>'
        shared = ('<w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/>' + text_run('beta')
                  + '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>'
                  + '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>')
        elsewhere = (f'<w:p>{text_run("Alpha ")}{anchored(0, "beta")}{text_run(" gamma.")}</w:p>'
                     f'<w:p>{anchored(1, "Other")}</w:p>')
        task = complex_task('docx', {'comments': [{'action': 'reply', 'part': 'word/document.xml', 'comment': 'Check this',
                                                   'author': 'Editor', 'text': 'Checked'}]})
        self.assert_verdicts(task, lambda: word_package(body(anchored(0, 'beta')), parent, '<w15:commentEx w15:paraId="10000001"/>'), {
            'reply without anchors': (lambda: word_package(body(anchored(0, 'beta')), parent + reply, threaded), True),
            'reply on the parent range': (lambda: word_package(body(shared), parent + reply, threaded), True),
            'reply anchors without threading': (lambda: word_package(body(shared), parent + reply, unthreaded), False),
            'reply anchored elsewhere': (lambda: word_package(elsewhere, parent + reply, threaded), False),
            'reply anchors replacing the parent anchors': (lambda: word_package(body(anchored(1, 'beta')), parent + reply, threaded), False),
        })

    def test_word_comment_reply_range_counts_tabs_breaks_and_paragraph_ends(self):
        parent = word_comment(0, 'Reviewer', 'Check this', '10000001')
        reply = word_comment(1, 'Editor', 'Checked', '10000002')
        threaded = '<w15:commentEx w15:paraId="10000001"/><w15:commentEx w15:paraId="10000002" w15:paraIdParent="10000001"/>'
        ends = ('<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>'
                '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>')
        def body(separator, before, after=''):
            return f'<w:p>{text_run("Alpha")}{before}{separator}{after}{text_run("beta")}{ends}{text_run(" gamma.")}</w:p>'
        both = '<w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/>'
        task = complex_task('docx', {'comments': [{'action': 'reply', 'part': 'word/document.xml', 'comment': 'Check this',
                                                   'author': 'Editor', 'text': 'Checked'}]})
        for name, separator in (('tab', '<w:r><w:tab/></w:r>'), ('line break', '<w:r><w:br/></w:r>'),
                                ('carriage return', '<w:r><w:cr/></w:r>')):
            source = f'<w:p>{text_run("Alpha")}{separator}{anchored(0, "beta")}{text_run(" gamma.")}</w:p>'
            self.assert_verdicts(task, lambda source=source: word_package(source, parent, '<w15:commentEx w15:paraId="10000001"/>'), {
                f'reply on the parent range after a {name}': (lambda separator=separator: word_package(
                    body(separator, '', both), parent + reply, threaded), True),
                f'reply starting before a {name}': (lambda separator=separator: word_package(
                    body(separator, '<w:commentRangeStart w:id="1"/>', '<w:commentRangeStart w:id="0"/>'), parent + reply, threaded), False),
            })
        source = f'<w:p>{text_run("Alpha")}</w:p><w:p>{anchored(0, "beta")}{text_run(" gamma.")}</w:p>'
        self.assert_verdicts(task, lambda: word_package(source, parent, '<w15:commentEx w15:paraId="10000001"/>'), {
            'reply starting in the previous paragraph': (lambda: word_package(
                f'<w:p>{text_run("Alpha")}<w:commentRangeStart w:id="1"/></w:p><w:p><w:commentRangeStart w:id="0"/>'
                f'{text_run("beta")}{ends}{text_run(" gamma.")}</w:p>', parent + reply, threaded), False),
        })

    def test_tracked_change_may_carry_its_boundary_space_on_either_side(self):
        def revision(kind, text):
            tag = 'delText' if kind == 'del' else 't'
            space = ' xml:space="preserve"' if text != text.strip() else ''
            return f'<w:{kind} w:id="1" w:author="Editor"><w:r><w:{tag}{space}>{text}</w:{tag}></w:r></w:{kind}>'
        def package(*items):
            return lambda: word_package(f'<w:p>{"".join(items)}</w:p>', '', '')
        insert = complex_task('docx', {'tracked_changes': [{'action': 'insert', 'part': 'word/document.xml', 'paragraph': 'Hire staff now.',
                                                            'after': 'Hire', 'text': ' and train', 'author': 'Editor'}]})
        self.assert_verdicts(insert, package(text_run('Hire staff now.')), {
            'space before the insertion': (package(text_run('Hire'), revision('ins', ' and train'), text_run(' staff now.')), True),
            'space after the insertion': (package(text_run('Hire '), revision('ins', 'and train '), text_run('staff now.')), True),
            'space dropped': (package(text_run('Hire '), revision('ins', 'and train'), text_run('staff now.')), False),
            'space doubled': (package(text_run('Hire'), revision('ins', 'and train '), text_run(' staff now.')), False),
        })
        delete = complex_task('docx', {'tracked_changes': [{'action': 'delete', 'part': 'word/document.xml', 'paragraph': 'Rule (No. 2) 2011',
                                                            'text': ' (No. 2)', 'author': 'Editor'}]})
        self.assert_verdicts(delete, package(text_run('Rule (No. 2) 2011')), {
            'space deleted before the text': (package(text_run('Rule'), revision('del', ' (No. 2)'), text_run(' 2011')), True),
            'space deleted after the text': (package(text_run('Rule '), revision('del', '(No. 2) '), text_run('2011')), True),
            'space deleted after text split across runs': (package(text_run('Rule '), '<w:del w:id="1" w:author="Editor"><w:r><w:delText>(No.</w:delText></w:r>'
                                                            '<w:r><w:delText xml:space="preserve"> 2) </w:delText></w:r></w:del>', text_run('2011')), True),
            'space kept twice': (package(text_run('Rule '), revision('del', '(No. 2)'), text_run(' 2011')), False),
        })
        insert['expected']['tracked_changes'][0].update(after='Hire ', text='and train ')
        self.assert_verdicts(insert, package(text_run('Hire staff now.')), {
            'requested trailing space moved before the insertion': (package(text_run('Hire'), revision('ins', ' and train'), text_run(' staff now.')), True),
            'requested trailing space dropped': (package(text_run('Hire'), revision('ins', ' and train'), text_run('staff now.')), False),
        })
        delete['expected']['tracked_changes'][0]['text'] = '(No. 2) '
        self.assert_verdicts(delete, package(text_run('Rule (No. 2) 2011')), {
            'requested trailing space deleted before the text': (package(text_run('Rule'), revision('del', ' (No. 2)'), text_run(' 2011')), True),
            'other text deleted': (package(text_run('Rule'), revision('del', ' (No. 2'), text_run(') 2011')), False),
        })

    def test_sort_may_add_the_hidden_filter_name_for_its_autofilter(self):
        task = complex_task('xlsx', {'autofilter_sort': [{'sheet': 'Data', 'ref': 'A1:B3', 'range': 'A2:B3',
                                                          'keys': [{'column': 'B', 'ascending': True}], 'fixed_rows': [1]}]})
        rows = [(1, 2), (10, 30), (20, 5)]
        sorted_rows = [(1, 2), (20, 5), (10, 30)]
        def filtered(names=''):
            return lambda: sheet_package(sorted_rows, 'A1:B3', names)
        def name(text, hidden=' hidden="1"'):
            return f'<definedName name="_xlnm._FilterDatabase" localSheetId="0"{hidden}>{text}</definedName>'
        self.assert_verdicts(task, lambda: sheet_package(rows), {
            'no filter name': (filtered(), True),
            'hidden filter name on the autofilter': (filtered(name('Data!$A$1:$B$3')), True),
            'quoted sheet in the filter name': (filtered(name("'Data'!$A$1:$B$3")), True),
            'filter name on another range': (filtered(name('Data!$A$1:$B$2')), False),
            'visible filter name': (filtered(name('Data!$A$1:$B$3', '')), False),
            'unsorted rows': (lambda: sheet_package(rows, 'A1:B3', name('Data!$A$1:$B$3')), False),
        })

    def test_slide_comment_reply_uses_the_powerpoint_parent_element(self):
        task = complex_task('pptx', {'comments': [{'slide': 1, 'comment': 'Check this', 'author': 'Editor', 'text': 'Checked'}]})
        owner = '<p:cmAuthor id="0" name="Reviewer" initials="R" lastIdx="2" clrIdx="0"/>'
        authors = owner + '<p:cmAuthor id="1" name="Editor" initials="E" lastIdx="1" clrIdx="1"/>'
        existing = slide_comment(0, 1, 'Check this') + slide_comment(0, 2, 'Other note')
        def reply(thread):
            return lambda: slide_package(authors, existing + slide_comment(1, 1, 'Checked', thread))
        self.assert_verdicts(task, lambda: slide_package(owner, existing), {
            'parentCm element': (reply('<p15:threadingInfo timeZoneBias="0"><p15:parentCm authorId="0" idx="1"/></p15:threadingInfo>'), True),
            'parentCm on another comment': (reply('<p15:threadingInfo timeZoneBias="0"><p15:parentCm authorId="0" idx="2"/></p15:threadingInfo>'), False),
            'parentCm on a missing comment': (reply('<p15:threadingInfo timeZoneBias="0"><p15:parentCm authorId="0" idx="9"/></p15:threadingInfo>'), False),
            'parent attributes PowerPoint does not write': (reply('<p15:threadingInfo parentCmId="1" parentAuthorId="0"/>'), False),
            'not a reply': (reply('<p15:threadingInfo timeZoneBias="0"/>'), False),
        })

    def test_slide_placeholder_on_a_new_layout_does_not_need_the_layout_prompt_flag(self):
        task = complex_task('pptx', {'layout': [{'slide': 1, 'name': 'Title and Content', 'placeholders': {
            'Title 1': {'layout_placeholder': 'Title 1', 'text': 'Heading'},
            'Subtitle 2': {'layout_placeholder': 'Content Placeholder 2', 'text': 'Review copy'}}}]})
        def moved(placeholder):
            return lambda: layout_slide_package(placeholder, 'slideLayout2.xml')
        self.assert_verdicts(task, lambda: layout_slide_package('<p:ph type="subTitle" idx="1"/>'), {
            'placeholder as PowerPoint writes it': (moved('<p:ph idx="1"/>'), True),
            'layout prompt flag copied': (moved('<p:ph idx="1" hasCustomPrompt="1"/>'), True),
            'placeholder on another index': (moved('<p:ph idx="2"/>'), False),
            'placeholder keeps its old type': (moved('<p:ph type="subTitle" idx="1"/>'), False),
        })

    def test_slide_table_frame_may_refit_its_grid_like_powerpoint(self):
        a, b, c = (1000, 'a'), (2000, 'b'), (3000, 'c')
        first, second = (400, '1'), (500, '2')
        other = table_frame(3, [(700, 'x'), (800, 'y')], [(300, '9')], (1500, 300), (100, 900))
        def deck(columns, rows, size, offset=(100, 200), untouched=other):
            return lambda: table_slide_package(table_frame(2, columns, rows, size, offset) + untouched)
        column = complex_task('pptx', {'table_structure': [{'part': 'ppt/slides/slide1.xml', 'table': 1, 'axis': 'column',
                                                            'action': 'delete', 'index': 2}]})
        moved_other = table_frame(3, [(700, 'x'), (800, 'y')], [(300, '9')], (1500, 400), (100, 900))
        self.assert_verdicts(column, deck([a, b, c], [first, second], (6000, 900)), {
            'frame left as it was': (deck([a, c], [first, second], (6000, 900)), True),
            'frame refit to the grid': (deck([a, c], [first, second], (4000, 900)), True),
            'frame taller than its rows': (deck([a, c], [first, second], (4000, 1200)), True),
            'frame width off the grid': (deck([a, c], [first, second], (4500, 900)), False),
            'frame shorter than its rows': (deck([a, c], [first, second], (4000, 800)), False),
            'frame moved': (deck([a, c], [first, second], (4000, 900), (150, 200)), False),
            'kept column width changed': (deck([a, (2000, 'c')], [first, second], (3000, 900)), False),
            'untouched table resized': (deck([a, c], [first, second], (4000, 900), untouched=moved_other), False),
        })
        row = complex_task('pptx', {'table_structure': [{'part': 'ppt/slides/slide1.xml', 'table': 1, 'axis': 'row',
                                                         'action': 'delete', 'index': 1}]})
        self.assert_verdicts(row, deck([a, b, c], [first, second], (6000, 900)), {
            'frame left as it was': (deck([a, b, c], [second], (6000, 900)), True),
            'frame refit to the rows': (deck([a, b, c], [second], (6000, 500)), True),
            'frame shorter than its rows': (deck([a, b, c], [second], (6000, 450)), False),
        })

    def test_chart_given_a_title_may_take_the_form_excel_writes(self):
        task = complex_task('xlsx', {'chart': [{'part': 'xl/charts/chart1.xml', 'title': 'Totals', 'series': [
            {'name': 'New', 'categories': 'Data!$A$1:$A$2', 'values': 'Data!$B$1:$B$2'}]}]})
        def titled(head):
            return lambda: chart_workbook(head, 'New')
        self.assert_verdicts(task, lambda: chart_workbook('<c:autoTitleDeleted val="1"/>'), {
            'title as the source flags left it': (titled(chart_title() + '<c:autoTitleDeleted val="1"/>'), True),
            'title as Excel writes it': (titled(chart_title('<c:overlay val="0"/>') + '<c:autoTitleDeleted val="0"/>'), True),
            'title without the deleted flag': (titled(chart_title('<c:overlay val="0"/>')), True),
            'title overlaying the plot': (titled(chart_title('<c:overlay val="1"/>') + '<c:autoTitleDeleted val="0"/>'), False),
            'wrong title text': (titled(chart_title().replace('Totals', 'Total') + '<c:autoTitleDeleted val="0"/>'), False),
        })

    def test_slide_reply_to_a_reply_may_thread_to_its_root_like_powerpoint(self):
        task = complex_task('pptx', {'comments': [{'slide': 1, 'comment': 'Second thought', 'author': 'Editor', 'text': 'Checked'}]})
        owner = '<p:cmAuthor id="0" name="Reviewer" initials="R" lastIdx="3" clrIdx="0"/>'
        authors = owner + '<p:cmAuthor id="1" name="Editor" initials="E" lastIdx="1" clrIdx="1"/>'
        def parent(index):
            return f'<p15:threadingInfo timeZoneBias="0"><p15:parentCm authorId="0" idx="{index}"/></p15:threadingInfo>'
        existing = (slide_comment(0, 1, 'Check this') + slide_comment(0, 2, 'Second thought', parent(1))
                    + slide_comment(0, 3, 'Other note'))
        def reply(thread):
            return lambda: slide_package(authors, existing + slide_comment(1, 1, 'Checked', thread))
        self.assert_verdicts(task, lambda: slide_package(owner, existing), {
            'reply on the thread root': (reply(parent(1)), True),
            'reply on the addressed reply': (reply(parent(2)), True),
            'reply on another thread': (reply(parent(3)), False),
            'reply on a missing comment': (reply(parent(9)), False),
            'not a reply': (reply('<p15:threadingInfo timeZoneBias="0"/>'), False),
        })

    def test_slide_reply_root_acceptance_resolves_authors_and_duplicate_reply_text(self):
        task = complex_task('pptx', {'comments': [{'slide': 1, 'comment': 'Second thought', 'author': 'Editor', 'text': 'Checked'}]})
        owner = '<p:cmAuthor id="0" name="Reviewer" initials="R" lastIdx="3" clrIdx="0"/>'
        editor = '<p:cmAuthor id="1" name="Editor" initials="E" lastIdx="1" clrIdx="1"/>'
        def parent(author, index):
            return f'<p15:threadingInfo timeZoneBias="0"><p15:parentCm authorId="{author}" idx="{index}"/></p15:threadingInfo>'
        def existing(author, duplicate):
            return (slide_comment(author, 1, 'Check this')
                    + slide_comment(author, 2, 'Second thought', parent(author, 1))
                    + slide_comment(author, 3, 'Checked' if duplicate else 'Other note'))
        for duplicate in (False, True):
            for renumber in (False, True):
                with self.subTest(duplicate=duplicate, renumber=renumber):
                    saved_owner, saved_editor = (1, 0) if renumber else (0, 1)
                    authors = (owner.replace('id="0"', f'id="{saved_owner}"')
                               + editor.replace('id="1"', f'id="{saved_editor}"'))
                    def reply(thread):
                        return lambda: slide_package(authors, existing(saved_owner, duplicate)
                                                     + slide_comment(saved_editor, 1, 'Checked', thread))
                    self.assert_verdicts(task, lambda: slide_package(owner, existing(0, duplicate)), {
                        'reply on the thread root': (reply(parent(saved_owner, 1)), True),
                        'reply on the addressed reply': (reply(parent(saved_owner, 2)), True),
                        'reply on another thread': (reply(parent(saved_owner, 3)), False),
                        'reply on a missing comment': (reply(parent(saved_owner, 9)), False),
                        'not a reply': (reply('<p15:threadingInfo timeZoneBias="0"/>'), False),
                    })


class InterruptTests(unittest.TestCase):
    def test_sigint_cancels_queued_runs_and_kills_agent_groups(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            fixture = root / 'fixture.xlsx'
            fixture.write_bytes(b'fixture')
            marker = root / 'started'
            worker = root / 'worker.py'
            worker.write_text('import os, subprocess, sys, time\nfrom pathlib import Path\n'
                              'child = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(60)"])\n'
                              f'with open({str(marker)!r}, "a") as stream: stream.write(str(os.getpid()) + " " + str(child.pid) + "\\n")\n'
                              'time.sleep(60)\n')
            harness = root / 'harness.py'
            harness.write_text(f'''import importlib.util, sys
from pathlib import Path
spec = importlib.util.spec_from_file_location('runner', {str(ROOT / 'run.py')!r})
r = importlib.util.module_from_spec(spec)
spec.loader.exec_module(r)
r.checker.fixture = lambda task: Path({str(fixture)!r})
r.checker.prepare_input = lambda task, data: data
r.checker.check = lambda task, output: None
r.python_ready = lambda python: None
r.shutil.which = lambda command: sys.executable
r.command_line = lambda *args: [sys.executable, {str(worker)!r}]
task = {{'id': 'interrupt', 'format': 'xlsx', 'input': 'input.xlsx', 'output': 'output.xlsx', 'expected': {{}}, 'prompt': 'test'}}
r.suite_tasks = lambda *args: [dict(task, id='interrupt-' + str(i)) for i in range(8)]
sys.argv = ['run.py', '--arm', 'generic', '--python', sys.executable, '--jobs', '2', {str(root / 'results')!r}]
sys.exit(r.main())
''')
            process = subprocess.Popen([sys.executable, str(harness)], start_new_session=True,
                                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            completed = False
            pids = []
            try:
                deadline = time.monotonic() + 10
                while time.monotonic() < deadline:
                    if marker.exists() and len(marker.read_text().splitlines()) == 2:
                        break
                    time.sleep(0.02)
                self.assertTrue(marker.exists())
                self.assertEqual(len(marker.read_text().splitlines()), 2)
                pids = [int(pid) for pid in marker.read_text().split()]
                process.send_signal(signal.SIGINT)
                try:
                    process.wait(timeout=5)
                    completed = True
                except subprocess.TimeoutExpired:
                    pass
                self.assertTrue(completed, 'interrupted batch waited for queued or running agents')
                self.assertNotEqual(process.returncode, 0)
                self.assertEqual(len(marker.read_text().splitlines()), 2)
                for pid in pids:
                    status = Path(f'/proc/{pid}/stat')
                    self.assertTrue(not status.exists() or status.read_text().split()[2] == 'Z', f'agent descendant {pid} survived')
            finally:
                if process.poll() is None:
                    os.killpg(process.pid, signal.SIGKILL)
                for pid in pids:
                    try:
                        os.kill(pid, signal.SIGKILL)
                    except ProcessLookupError:
                        pass
                process.wait()


if __name__ == '__main__':
    unittest.main()
