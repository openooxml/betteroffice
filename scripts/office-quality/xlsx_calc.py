import io
import math
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET
import xml.sax
from xml.sax.saxutils import XMLGenerator
from xml.sax.xmlreader import AttributesImpl

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
MAX_EXPANDED = 64 * 1024 * 1024
MAX_RESULTS = 250_000


def package(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        entries = archive.infolist()
        if len(entries) > 10000 or sum(entry.file_size for entry in entries) > MAX_EXPANDED:
            raise ValueError('Workbook exceeds calculation benchmark size limit')
        if len({entry.filename for entry in entries}) != len(entries):
            raise ValueError('Duplicate workbook parts')
        return {entry.filename: archive.read(entry) for entry in entries}


def sheets(parts, worksheets_only=False):
    workbook = ET.fromstring(parts['xl/workbook.xml'])
    relationships = ET.fromstring(parts['xl/_rels/workbook.xml.rels'])
    targets = {r.get('Id'): r.get('Target') for r in relationships if r.get('TargetMode') != 'External'}
    worksheet_ids = {r.get('Id') for r in relationships
                     if r.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
                     and r.get('TargetMode') != 'External'}
    result = []
    for sheet in workbook.findall('s:sheets/s:sheet', NS):
        if worksheets_only and sheet.get(REL) not in worksheet_ids:
            continue
        target = targets[sheet.get(REL)]
        path = posixpath.normpath(target.lstrip('/') if target.startswith('/') else 'xl/' + target)
        if not path.startswith('xl/') or path not in parts:
            raise ValueError('Invalid worksheet relationship')
        result.append((sheet.get('name'), path, ET.fromstring(parts[path])))
    return result


def address(value):
    match = re.fullmatch(r'([A-Z]{1,3})([1-9][0-9]*)', value)
    if not match:
        raise ValueError('Invalid cell address')
    column = 0
    for char in match[1]:
        column = column * 26 + ord(char) - 64
    row = int(match[2])
    if row > 1048576 or column > 16384:
        raise ValueError('Cell outside worksheet')
    return row - 1, column - 1


def a1(row, column):
    name = ''
    column += 1
    while column:
        column, remainder = divmod(column - 1, 26)
        name = chr(65 + remainder) + name
    return name + str(row + 1)


def formula_targets(parts):
    result = []
    formulas = 0
    for index, (_, _, sheet) in enumerate(sheets(parts)):
        targets = set()
        for cell in sheet.findall('s:sheetData/s:row/s:c', NS):
            formula = cell.find('s:f', NS)
            if formula is None:
                continue
            formulas += 1
            targets.add(address(cell.get('r')))
            if formula.get('t') == 'array' and formula.get('ref'):
                bounds = formula.get('ref').split(':')
                top, left = address(bounds[0])
                bottom, right = address(bounds[-1])
                if bottom < top or right < left or (bottom-top+1)*(right-left+1) > MAX_RESULTS:
                    raise ValueError('Array formula exceeds result limit')
                targets.update((r,c) for r in range(top,bottom+1) for c in range(left,right+1))
        result.extend(dict(sheet=index,cell=a1(row,col)) for row,col in sorted(targets))
        if len(result) > MAX_RESULTS:
            raise ValueError('Workbook exceeds formula result limit')
    return result, formulas


def text_value(node):
    text = ''.join(t.text or '' for t in node.findall('.//s:t', NS))
    return re.sub(r'_x([0-9a-fA-F]{4})_', lambda m: chr(int(m[1],16)), text)


def results(parts, targets):
    shared = []
    if 'xl/sharedStrings.xml' in parts:
        shared = [text_value(item) for item in ET.fromstring(parts['xl/sharedStrings.xml'])]
    cells = [{cell.get('r'): cell for cell in sheet.findall('s:sheetData/s:row/s:c', NS)} for _,_,sheet in sheets(parts)]
    output = []
    for target in targets:
        cell = cells[target['sheet']].get(target['cell'])
        value = {'kind': 'empty'}
        if cell is not None:
            raw = cell.find('s:v', NS)
            kind = cell.get('t', 'n')
            if kind == 'inlineStr':
                value = {'kind':'text','value':text_value(cell)}
            elif raw is not None:
                raw = raw.text or ''
                if kind == 's':
                    value = {'kind':'text','value':shared[int(raw)]}
                elif kind == 'str':
                    value = {'kind':'text','value':re.sub(r'_x([0-9a-fA-F]{4})_',lambda m:chr(int(m[1],16)),raw)}
                elif kind == 'b' and raw in {'0','1'}:
                    value = {'kind':'bool','value':raw == '1'}
                elif kind == 'e':
                    value = {'kind':'error','value':raw}
                elif kind == 'n' and raw:
                    number = float(raw)
                    if not math.isfinite(number):
                        raise ValueError('Nonfinite formula result')
                    value = {'kind':'number','value':number}
                elif raw:
                    raise ValueError('Unsupported result type: ' + kind)
        output.append(dict(target, value=value))
    return output


class UncachedSheet(XMLGenerator):
    def __init__(self, output, targets):
        super().__init__(output, encoding='utf-8', short_empty_elements=True)
        self.targets = targets
        self.clear = False
        self.skipped = 0

    def startElement(self, name, attrs):
        local = name.split(':')[-1]
        if self.skipped:
            self.skipped += 1
            return
        if local == 'c':
            self.clear = attrs.get('r') in self.targets
            if self.clear:
                attrs = AttributesImpl({k:v for k,v in attrs.items() if k != 't'})
        if self.clear and local in {'v','is'}:
            self.skipped = 1
            return
        super().startElement(name, attrs)

    def endElement(self, name):
        if self.skipped:
            self.skipped -= 1
            return
        super().endElement(name)
        if name.split(':')[-1] == 'c':
            self.clear = False

    def characters(self, content):
        if not self.skipped:
            super().characters(content)


class AutomaticCalculation(XMLGenerator):
    def __init__(self, output):
        super().__init__(output, encoding='utf-8', short_empty_elements=True)
        self.seen = False

    def startElement(self, name, attrs):
        if name.split(':')[-1] == 'calcPr':
            self.seen = True
            attrs = AttributesImpl({**dict(attrs.items()),'calcMode':'auto','fullCalcOnLoad':'1','forceFullCalc':'1','calcId':'0'})
        super().startElement(name,attrs)

    def endElement(self, name):
        if name.split(':')[-1] == 'workbook' and not self.seen:
            prefix = name.rsplit(':',1)[0]+':' if ':' in name else ''
            self.startElement(prefix+'calcPr',AttributesImpl({}))
            super().endElement(prefix+'calcPr')
        super().endElement(name)


def uncached_workbook(data, targets):
    parts = package(data)
    names = sheets(parts)
    by_part = {path:{t['cell'] for t in targets if t['sheet'] == index} for index,(_,path,_) in enumerate(names)}
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as original, zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as archive:
        for info in original.infolist():
            content = parts[info.filename]
            if by_part.get(info.filename):
                cleared = io.BytesIO()
                xml.sax.parse(io.BytesIO(content), UncachedSheet(cleared,by_part[info.filename]))
                content = cleared.getvalue()
            if info.filename == 'xl/workbook.xml':
                automatic = io.BytesIO()
                xml.sax.parse(io.BytesIO(content), AutomaticCalculation(automatic))
                content = automatic.getvalue()
            archive.writestr(info,content)
    return output.getvalue()
