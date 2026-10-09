import argparse
import copy
import json
import posixpath
import re
import xml.etree.ElementTree as ET
from decimal import Decimal
from pathlib import Path
from urllib.parse import unquote
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parent
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
BUILTIN_FORMATS = {0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
                   9: '0%', 10: '0.00%', 14: 'mm-dd-yy', 49: '@'}


def tasks():
    return json.loads((ROOT / 'tasks.json').read_text())


def require(condition, message):
    if not condition:
        raise ValueError(message)


def xml(archive, path):
    return ET.fromstring(archive.read(path))


def canonical(node):
    if node is None:
        return None
    text = node.text if node.text and node.text.strip() else None
    return (node.tag, tuple(sorted(node.attrib.items())), text,
            tuple(canonical(child) for child in node))


def relpath(part):
    folder, name = posixpath.split(part)
    return posixpath.join(folder, '_rels', name + '.rels') if part else '_rels/.rels'


def relations(archive, part):
    path = relpath(part)
    if path not in archive.namelist():
        return {}
    result = {}
    for rel in xml(archive, path):
        key = rel.attrib['Id']
        require(key not in result, f'duplicate relationship ID in {path}: {key}')
        target = rel.attrib['Target']
        external = rel.get('TargetMode') == 'External'
        target = target if external else posixpath.normpath(posixpath.join(posixpath.dirname(part), unquote(target)))
        result[key] = {'type': rel.attrib['Type'].rsplit('/', 1)[-1],
                       'target': target.lstrip('/') if not external else target, 'external': external}
    return result


def related(archive, part, kind):
    matches = [rel['target'] for rel in relations(archive, part).values() if rel['type'] == kind and not rel['external']]
    require(len(matches) <= 1, f'multiple {kind} relationships from {part}')
    return matches[0] if matches else None


def validate_package(archive):
    names = archive.namelist()
    require(len(names) == len(set(names)), 'duplicate ZIP members')
    require(archive.testzip() is None, 'invalid ZIP checksum')
    for name in names:
        if name.endswith(('.xml', '.rels')):
            xml(archive, name)
        if name.endswith('.rels'):
            folder, filename = posixpath.split(name)
            part = posixpath.join(posixpath.dirname(folder), filename[:-5]) if name != '_rels/.rels' else ''
            for rel in relations(archive, part).values():
                require(rel['external'] or rel['target'] in names,
                        f'missing relationship target: {name} -> {rel["target"]}')
    for item in xml(archive, '[Content_Types].xml'):
        if item.tag.endswith('Override'):
            require(item.attrib['PartName'].lstrip('/') in names, f'missing content type part: {item.attrib["PartName"]}')


def boolean(node):
    return node is not None and node.get('val', '1') not in ('0', 'false')


def styles(archive):
    part = related(archive, 'xl/workbook.xml', 'styles')
    require(part is not None, 'workbook styles are missing')
    root = xml(archive, part)
    formats = dict(BUILTIN_FORMATS)
    formats.update({int(n.attrib['numFmtId']): n.attrib['formatCode'] for n in root.findall('s:numFmts/s:numFmt', NS)})
    fonts = root.findall('s:fonts/s:font', NS)
    fills = root.findall('s:fills/s:fill', NS)
    borders = root.findall('s:borders/s:border', NS)
    result = []
    for xf in root.findall('s:cellXfs/s:xf', NS):
        font = copy.deepcopy(fonts[int(xf.get('fontId', '0'))])
        bold, italic = boolean(font.find('s:b', NS)), boolean(font.find('s:i', NS))
        for tag in ('b', 'i'):
            for node in font.findall(f's:{tag}', NS):
                font.remove(node)
        fill = fills[int(xf.get('fillId', '0'))].find('s:patternFill', NS)
        color = fill.find('s:fgColor', NS) if fill is not None else None
        fmt = int(xf.get('numFmtId', '0'))
        require(fmt in formats, f'unknown number format {fmt}')
        result.append({'number_format': formats[fmt], 'bold': bold, 'italic': italic,
                       'fill_pattern': fill.get('patternType', 'none') if fill is not None else 'none',
                       'fill': color.get('rgb', '').upper() if color is not None else None,
                       'fill_other': tuple(sorted((k, v) for k, v in color.attrib.items() if k != 'rgb')) if color is not None else (),
                       'font': canonical(font), 'border': canonical(borders[int(xf.get('borderId', '0'))]),
                       'alignment': canonical(xf.find('s:alignment', NS)),
                       'protection': canonical(xf.find('s:protection', NS))})
    require(bool(result), 'no cell styles')
    return result


def workbook(archive):
    cell_styles = styles(archive)
    strings_part = related(archive, 'xl/workbook.xml', 'sharedStrings')
    strings = [''.join(n.itertext()) for n in xml(archive, strings_part).findall('s:si', NS)] if strings_part else []
    links = relations(archive, 'xl/workbook.xml')
    result = []
    for entry in xml(archive, 'xl/workbook.xml').findall('s:sheets/s:sheet', NS):
        link = links[entry.attrib[f'{{{NS["r"]}}}id']]
        require(link['type'] == 'worksheet' and not link['external'], 'sheet must reference a worksheet')
        sheet = xml(archive, link['target'])
        cells = {}
        for node in sheet.findall('s:sheetData/s:row/s:c', NS):
            address = node.attrib['r']
            require(address not in cells, f'duplicate cell {address}')
            kind, value = node.get('t', 'n'), node.findtext('s:v', namespaces=NS)
            formula = node.find('s:f', NS)
            require(formula is None or formula.get('t', 'normal') == 'normal', 'shared or array formula cannot be checked')
            if kind == 'inlineStr':
                value = ''.join(n.text or '' for n in node.findall('s:is//s:t', NS))
            elif kind == 's':
                value = strings[int(value)]
            elif kind == 'b':
                value = value == '1'
            elif kind == 'n' and value is not None:
                value = Decimal(value)
                require(value.is_finite(), f'non-finite value in {address}')
            elif kind == 'e':
                value = {'error': value}
            style = copy.deepcopy(cell_styles[int(node.get('s', '0'))])
            record = {'value': value, 'formula': formula.text if formula is not None else None, 'style': style}
            if value is not None or formula is not None or style != cell_styles[0]:
                cells[address] = record
        metadata = [canonical(n) for n in sheet if n.tag not in (f'{{{NS["s"]}}}sheetData', f'{{{NS["s"]}}}dimension')]
        rows = {n.attrib['r']: {k: v for k, v in n.attrib.items() if k not in ('r', 'spans')}
                for n in sheet.findall('s:sheetData/s:row', NS) if set(n.attrib) - {'r', 'spans'}}
        result.append({'name': entry.attrib['name'], 'state': entry.get('state', 'visible'),
                       'cells': cells, 'metadata': metadata, 'row_properties': rows})
    require(len({sheet['name'] for sheet in result}) == len(result), 'duplicate worksheet names')
    return {'sheets': result, 'default_style': cell_styles[0]}


def column_number(letters):
    number = 0
    for char in letters:
        number = number * 26 + ord(char) - 64
    return number


def column_letters(number):
    text = ''
    while number:
        number, digit = divmod(number - 1, 26)
        text = chr(65 + digit) + text
    return text


def expected_workbook(source, spec):
    expected = copy.deepcopy(source)
    for sheet in expected['sheets']:
        name = sheet['name']
        if name in spec.get('insert_columns', {}):
            at = spec['insert_columns'][name]
            moved = {}
            for address, record in sheet['cells'].items():
                letters, row = re.fullmatch(r'([A-Z]+)([0-9]+)', address).groups()
                col = column_number(letters)
                moved[f'{column_letters(col + (col >= at))}{row}'] = record
            sheet['cells'] = moved
        if name in spec.get('rows', {}):
            original = copy.deepcopy(sheet['cells'])
            for destination, origin in spec['rows'][name].items():
                for address in list(sheet['cells']):
                    if re.search(r'[0-9]+$', address).group() == destination:
                        del sheet['cells'][address]
                for address, record in original.items():
                    if re.search(r'[0-9]+$', address).group() == str(origin):
                        sheet['cells'][re.sub(r'[0-9]+$', destination, address)] = copy.deepcopy(record)
        sheet['name'] = spec.get('rename', {}).get(name, name)
        for address, patch in spec.get('cells', {}).get(sheet['name'], {}).items():
            record = sheet['cells'].setdefault(address, {'value': None, 'formula': None, 'style': copy.deepcopy(source['default_style'])})
            for key, value in patch.items():
                if key == 'style':
                    record['style'].update(value)
                else:
                    record[key] = Decimal(str(value)) if isinstance(value, (int, float)) and not isinstance(value, bool) else value
    if 'sheet_order' in spec:
        by_name = {sheet['name']: sheet for sheet in expected['sheets']}
        expected['sheets'] = [by_name[name] for name in spec['sheet_order']]
    return expected


def paragraph(node, inherited_bullet=False):
    props = node.find('a:pPr', NS)
    none = props is not None and props.find('a:buNone', NS) is not None
    explicit_bullet = props is not None and any(props.find(f'a:{kind}', NS) is not None for kind in ('buChar', 'buAutoNum', 'buBlip'))
    bullet = not none and (inherited_bullet or explicit_bullet)
    runs = []
    for run in node:
        if run.tag == f'{{{NS["a"]}}}br':
            runs.append({'text': '\n', 'style': canonical(run.find('a:rPr', NS))})
        elif run.tag in (f'{{{NS["a"]}}}r', f'{{{NS["a"]}}}fld'):
            text = run.findtext('a:t', default='', namespaces=NS)
            style = canonical(run.find('a:rPr', NS))
            if runs and runs[-1]['style'] == style:
                runs[-1]['text'] += text
            elif text:
                runs.append({'text': text, 'style': style})
    return {'text': ''.join(run['text'] for run in runs), 'bullet': bool(bullet), 'props': canonical(props), 'runs': runs}


def slide_notes(archive, part):
    target = related(archive, part, 'notesSlide')
    if not target:
        return []
    paragraphs = []
    for shape in xml(archive, target).findall('p:cSld/p:spTree/p:sp', NS):
        placeholder = shape.find('p:nvSpPr/p:nvPr/p:ph', NS)
        if placeholder is None or placeholder.get('type') == 'body':
            paragraphs.extend(paragraph(n)['text'] for n in shape.findall('p:txBody/a:p', NS))
    return paragraphs


def presentation(archive):
    root = xml(archive, 'ppt/presentation.xml')
    links = relations(archive, 'ppt/presentation.xml')
    slides, seen = [], set()
    for entry in root.findall('p:sldIdLst/p:sldId', NS):
        identity = entry.attrib['id']
        target = links[entry.attrib[f'{{{NS["r"]}}}id']]
        require(target['type'] == 'slide' and not target['external'], 'slide must reference a slide part')
        part = target['target']
        require(identity not in seen and part not in seen, 'duplicate presentation slide')
        seen.update((identity, part))
        slide = xml(archive, part)
        layout_part = related(archive, part, 'slideLayout')
        master_part = related(archive, layout_part, 'slideMaster') if layout_part else None
        master = xml(archive, master_part) if master_part else None
        shapes, shape_ids = [], set()
        tree = slide.find('p:cSld/p:spTree', NS)
        require(tree is not None, 'slide has no shape tree')
        for node in tree:
            if node.tag in (f'{{{NS["p"]}}}nvGrpSpPr', f'{{{NS["p"]}}}grpSpPr'):
                continue
            require(node.tag == f'{{{NS["p"]}}}sp', 'unexpected non-text shape')
            identity = node.find('p:nvSpPr/p:cNvPr', NS)
            require(identity is not None and identity.attrib['id'] not in shape_ids, 'duplicate or missing shape identity')
            shape_ids.add(identity.attrib['id'])
            props = copy.deepcopy(node.find('p:spPr', NS))
            transform = props.find('a:xfrm', NS) if props is not None else None
            rect = None
            if transform is not None:
                off, ext = transform.find('a:off', NS), transform.find('a:ext', NS)
                rect = [int(off.attrib['x']), int(off.attrib['y']), int(ext.attrib['cx']), int(ext.attrib['cy'])]
                props.remove(transform)
            placeholder = node.find('p:nvSpPr/p:nvPr/p:ph', NS)
            inherited = master is not None and placeholder is not None and placeholder.get('type') == 'body' and master.find('p:txStyles/p:bodyStyle/a:lvl1pPr/a:buChar', NS) is not None
            shapes.append({'name': identity.get('name', ''), 'rect': rect, 'properties': canonical(props),
                           'transform_properties': dict(transform.attrib) if transform is not None else {},
                           'placeholder': canonical(placeholder), 'body_properties': canonical(node.find('p:txBody/a:bodyPr', NS)),
                           'paragraphs': [paragraph(n, inherited) for n in node.findall('p:txBody/a:p', NS)]})
        slides.append({'shapes': shapes, 'notes': slide_notes(archive, part), 'layout': layout_part,
                       'background': canonical(slide.find('p:cSld/p:bg', NS)),
                       'group_properties': [canonical(node) for node in tree if node.tag in (f'{{{NS["p"]}}}nvGrpSpPr', f'{{{NS["p"]}}}grpSpPr')],
                       'slide_properties': [canonical(node) for node in slide if node.tag != f'{{{NS["p"]}}}cSld']})
    return {'slides': slides, 'size': dict(root.find('p:sldSz', NS).attrib),
            'assets': {name: canonical(xml(archive, name)) for name in archive.namelist()
                       if name.endswith('.xml') and name.startswith(('ppt/slideLayouts/', 'ppt/slideMasters/', 'ppt/notesMasters/', 'ppt/theme/'))}}


def replace_paragraph_text(item, text):
    require(len(item['runs']) == 1, 'fixture replacement must have one styled run')
    item['text'] = text
    item['runs'][0]['text'] = text


def expected_presentation(source, spec):
    expected = copy.deepcopy(source)
    for i, slide in enumerate(expected['slides'], 1):
        for shape in slide['shapes']:
            if 'replace' in spec:
                for item in shape['paragraphs']:
                    replace_paragraph_text(item, item['text'].replace(spec['replace']['old'], spec['replace']['new']))
            texts = spec.get('changes', {}).get(str(i), {}).get(shape['name'])
            if texts is not None:
                require(len(texts) == len(shape['paragraphs']), 'invalid paragraph expectation')
                for item, text in zip(shape['paragraphs'], texts):
                    replace_paragraph_text(item, text)
        if str(i) in spec.get('notes', {}):
            slide['notes'] = spec['notes'][str(i)].split('\n')
        slide['shapes'].extend({'requested': request} for request in spec.get('add_shapes', {}).get(str(i), []))
    if 'slide_order' in spec:
        expected['slides'] = [expected['slides'][item - 1] if isinstance(item, int) else {'requested': item} for item in spec['slide_order']]
    return expected


def compare(expected, actual, path='document'):
    if isinstance(expected, dict):
        require(isinstance(actual, dict) and expected.keys() == actual.keys(), f'{path}: fields differ')
        for key in expected:
            compare(expected[key], actual[key], f'{path}.{key}')
    elif isinstance(expected, (list, tuple)):
        require(isinstance(actual, (list, tuple)) and len(expected) == len(actual), f'{path}: count differs ({len(expected)} expected, {len(actual) if isinstance(actual, (list, tuple)) else "invalid"} found)')
        for i, (before, after) in enumerate(zip(expected, actual)):
            compare(before, after, f'{path}[{i + 1}]')
    else:
        require(expected == actual and isinstance(expected, bool) == isinstance(actual, bool), f'{path}: expected {expected!r}, found {actual!r}')


def check_requested_shape(request, actual, path):
    compare(request['texts'], [item['text'] for item in actual['paragraphs']], path + '.text')
    if 'bullets' in request:
        compare(request['bullets'], [item['bullet'] for item in actual['paragraphs']], path + '.bullets')
    if 'rect' in request:
        compare(request['rect'], actual['rect'], path + '.position')
    for item in actual['paragraphs']:
        for run in item['runs']:
            style = run['style']
            if 'font_size' in request:
                require(style is not None and dict(style[1]).get('sz') == str(request['font_size']), path + ': font size differs')
            if 'color' in request:
                require(style is not None and any(child[0] == f'{{{NS["a"]}}}solidFill' and
                        any(color[0] == f'{{{NS["a"]}}}srgbClr' and dict(color[1]).get('val', '').upper() == request['color'] for color in child[3]) for child in style[3]), path + ': text color differs')


def check_presentation(expected, actual):
    compare(expected['size'], actual['size'], 'slide size')
    compare(expected['assets'], actual['assets'], 'layouts and masters')
    require(len(expected['slides']) == len(actual['slides']), 'slide count differs')
    for i, (before, after) in enumerate(zip(expected['slides'], actual['slides']), 1):
        path = f'slide {i}'
        if 'requested' in before:
            request = before['requested']
            compare(request['notes'].split('\n') if request.get('notes') else [], after['notes'], path + '.notes')
            require(len(request['shapes']) == len(after['shapes']), path + ': shape count differs')
            for j, (shape, found) in enumerate(zip(request['shapes'], after['shapes']), 1):
                check_requested_shape(shape, found, f'{path}.shape {j}')
        else:
            for key in ('notes', 'layout', 'background', 'group_properties', 'slide_properties'):
                compare(before[key], after[key], f'{path}.{key}')
            require(len(before['shapes']) == len(after['shapes']), path + ': shape count differs')
            for j, (shape, found) in enumerate(zip(before['shapes'], after['shapes']), 1):
                if 'requested' in shape:
                    check_requested_shape(shape['requested'], found, f'{path}.shape {j}')
                else:
                    compare(shape, found, f'{path}.shape {j}')


def check(task, output):
    output = Path(output)
    require(output.is_file(), f'missing output file: {output.name}')
    with ZipFile(ROOT / task['fixture']) as source, ZipFile(output) as revised:
        validate_package(revised)
        if task['format'] == 'xlsx':
            compare(expected_workbook(workbook(source), task['expected']), workbook(revised))
        else:
            check_presentation(expected_presentation(presentation(source), task['expected']), presentation(revised))
    return 'exact requested content, formulas, formatting, order, notes, and preserved content verified'


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('task', choices=[task['id'] for task in tasks()])
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    task = next(task for task in tasks() if task['id'] == args.task)
    try:
        print(f'PASS: {task["id"]}: {check(task, args.output)}')
    except Exception as error:
        print(f'FAIL: {task["id"]}: {error}')
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
