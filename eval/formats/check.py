import argparse
import copy
import hashlib
import json
import math
import os
import posixpath
import re
import subprocess
import uuid
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
      'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
      'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
      'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
BUILTIN_FORMATS = {0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
                   9: '0%', 10: '0.00%', 14: 'mm-dd-yy', 49: '@'}
POOLS = ('numFmtId', 'fontId', 'fillId', 'borderId')
FACETS = {'applyNumberFormat': ('number_format',), 'applyFont': ('bold', 'italic', 'font'),
          'applyFill': ('fill_pattern', 'fill', 'fill_other', 'fill_background', 'fill_rest'),
          'applyBorder': ('border',), 'applyAlignment': ('alignment',), 'applyProtection': ('protection',)}
XF_FACET_CHILDREN = (f'{{{NS["s"]}}}alignment', f'{{{NS["s"]}}}protection')
ORDER_FREE_CHILDREN = {f'{{{NS["s"]}}}{name}' for name in ('font', 'rPr')}
NOTES_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'
OWNER = '<owning slide>'
CREATED = 'notes part created for the requested notes'
SAVE_TIME = '{http://purl.org/dc/terms/}modified'
PARAGRAPH_IDS = tuple(f'{{{NS["w14"]}}}{name}' for name in ('paraId', 'textId'))
TRACKED = {f'{{{NS["w"]}}}{name}' for name in ('ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange', 'sectPrChange',
                                         'tblPrChange', 'tblPrExChange', 'tblGridChange', 'trPrChange', 'tcPrChange',
                                         'cellIns', 'cellDel', 'cellMerge', 'numberingChange')}


def tasks(suite='v2'):
    require(suite in ('v2', 'v3'), f'unknown suite: {suite}')
    path = ROOT / 'v3' / 'tasks.json' if suite == 'v3' else ROOT / 'tasks.json'
    return [dict(task, suite=suite) for task in json.loads(path.read_text())]


def verified(data, entry, name):
    if len(data) != entry['bytes'] or hashlib.sha256(data).hexdigest() != entry['sha256']:
        raise ValueError(f'fixture {name} does not match fixtures.json')
    return data


def fixture(task):
    manifest = json.loads((ROOT / ('v3/fixtures.json' if task.get('suite', 'v2') == 'v3' else 'fixtures.json')).read_text())
    name = task['fixture']
    entry = manifest['fixtures'][name]
    local = os.environ.get('FORMAT_EVAL_FIXTURES', '').strip()
    if local:
        path = Path(local).expanduser().resolve() / name
        verified(path.read_bytes(), entry, name)
        return path
    cache = Path(os.environ.get('QUALITY_ASSET_CACHE', '').strip() or Path.home() / '.cache' / 'betteroffice' / 'eval-assets')
    path = cache / f"{entry['sha256']}-{entry['bytes']}"
    if path.is_file():
        try:
            verified(path.read_bytes(), entry, name)
            return path
        except ValueError:
            path.unlink()
    cache.mkdir(parents=True, exist_ok=True)
    staging = cache / f'.tmp-{uuid.uuid4()}'
    try:
        subprocess.run(['curl', '-fsSL', '--retry', '3', '--max-time', '120', '--max-filesize', str(entry['bytes']),
                        '-o', str(staging), f"{manifest['origin'].rstrip('/')}/{name}"], check=True)
        verified(staging.read_bytes(), entry, name)
        os.replace(staging, path)
    finally:
        staging.unlink(missing_ok=True)
    return path


def require(condition, message):
    if not condition:
        raise ValueError(message)


def xml(archive, path):
    return ET.fromstring(archive.read(path))


def canonical(node):
    if node is None:
        return None
    text = node.text if node.text and node.text.strip() else None
    children = [canonical(child) for child in node]
    if node.tag in ORDER_FREE_CHILDREN:
        children.sort(key=lambda child: child[0])
    return (node.tag, tuple(sorted(node.attrib.items())), text, tuple(children))


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


def resolved_xml(archive, part, node=None):
    node = copy.deepcopy(node if node is not None else xml(archive, part))
    links = relations(archive, part)
    for item in node.iter():
        for key, value in list(item.attrib.items()):
            if key.startswith(f'{{{NS["r"]}}}') and value:
                require(value in links, f'missing relationship ID in {part}: {value}')
                item.set(key, json.dumps(links[value], sort_keys=True))
    return node


def relationship_state(archive, part, dynamic=(), owner=None):
    path = relpath(part)
    if path not in archive.namelist():
        return ()
    links = relations(archive, part)
    records = []
    for item in xml(archive, path):
        link = links[item.attrib['Id']]
        if link['target'] == 'xl/calcChain.xml' and link['type'] == 'calcChain':
            continue
        if link['type'] in dynamic:
            continue
        attributes = {key: value for key, value in item.attrib.items() if key not in ('Id', 'Target')}
        owned = owner is not None and link['type'] == 'slide' and not link['external'] and link['target'] == owner
        attributes['Target'] = OWNER if owned else link['target']
        records.append(tuple(sorted(attributes.items())))
    return tuple(sorted(records))


def relationship_properties(archive, part, identity):
    item = next(item for item in xml(archive, relpath(part)) if item.get('Id') == identity)
    return {key: value for key, value in item.attrib.items() if key not in ('Id', 'Target')}


def content_type(archive, part):
    root = xml(archive, '[Content_Types].xml')
    overrides = [item.attrib['ContentType'] for item in root if item.get('PartName', '').lstrip('/') == part]
    require(len(overrides) <= 1, f'duplicate content type for {part}')
    if overrides:
        return overrides[0]
    extension = part.rsplit('.', 1)[-1]
    defaults = [item.attrib['ContentType'] for item in root if item.get('Extension') == extension]
    require(len(defaults) == 1, f'missing or duplicate default content type for {part}')
    return defaults[0]


def package_state(archive, modeled, dynamic_links):
    result = {}
    for name in archive.namelist():
        if name.endswith('/') or name in modeled or name == 'xl/calcChain.xml':
            continue
        if name.endswith('.rels'):
            folder, filename = posixpath.split(name)
            part = posixpath.join(posixpath.dirname(folder), filename[:-5]) if name != '_rels/.rels' else ''
            records = relationship_state(archive, part, dynamic_links.get(part, ()))
            if records:
                result[name] = records
        elif name == '[Content_Types].xml':
            records = []
            for item in xml(archive, name):
                if item.get('PartName', '').lstrip('/') in modeled | {'xl/calcChain.xml'}:
                    continue
                records.append(canonical(item))
            result[name] = tuple(sorted(records))
        elif name.endswith('.xml'):
            result[name] = canonical(resolved_xml(archive, name))
        else:
            result[name] = hashlib.sha256(archive.read(name)).hexdigest()
    return result


def workbook_state(archive, root):
    state = resolved_xml(archive, 'xl/workbook.xml', root)
    listing = state.find('s:sheets', NS)
    names = [sheet.attrib['name'] for sheet in listing]
    listing[:] = []
    for view in state.findall('s:bookViews/s:workbookView', NS):
        for key in ('activeTab', 'firstSheet'):
            index = int(view.get(key, '0'))
            require(0 <= index < len(names), f'workbookView {key} is out of range')
            view.set(key, names[index])
    for item in state.findall('s:definedNames/s:definedName', NS):
        if 'localSheetId' in item.attrib:
            index = int(item.attrib['localSheetId'])
            require(0 <= index < len(names), 'definedName localSheetId is out of range')
            item.set('localSheetId', names[index])
        item.text = rename_formula(item.text, {name: name for name in names})
    return canonical(state)


def selection_state(root, names, sheet_views):
    tabs = [int(view.get('activeTab', '0')) for view in root.findall('s:bookViews/s:workbookView', NS)] or [0]
    state = {view: set() for view in range(len(tabs))}
    for name, views in zip(names, sheet_views):
        for view, selected in views:
            state.setdefault(view, set())
            if selected:
                state[view].add(name)
    for view, tab in enumerate(tabs):
        require(0 <= tab < len(names), 'workbookView activeTab is out of range')
        state[view].add(names[tab])
    return {str(view): sorted(selected) for view, selected in sorted(state.items())}


def rename_formula(formula, renames):
    def reference(match):
        token = match.group()[:-1]
        name = token[1:-1].replace("''", "'") if token.startswith("'") else token
        if name not in renames:
            return match.group()
        return "'" + renames[name].replace("'", "''") + "'!"

    pieces = re.split(r'("(?:[^"]|"")*")', formula or '')
    for i in range(0, len(pieces), 2):
        pieces[i] = re.sub(r"(?<![\w.\]])(?:'(?:[^']|'')+'|[A-Za-z_\\][\w.\\]*)!", reference, pieces[i])
    return ''.join(pieces) if formula is not None else None


def expected_workbook_state(state, spec):
    tag, attributes, text, children = state
    attributes = dict(attributes)
    renames = spec.get('rename', {})
    if tag == f'{{{NS["s"]}}}workbookView':
        for key in ('activeTab', 'firstSheet'):
            if key in attributes:
                attributes[key] = renames.get(attributes[key], attributes[key])
    elif tag == f'{{{NS["s"]}}}definedName':
        if 'localSheetId' in attributes:
            attributes['localSheetId'] = renames.get(attributes['localSheetId'], attributes['localSheetId'])
        text = rename_formula(text, renames)
    return (tag, tuple(sorted(attributes.items())), text,
            tuple(expected_workbook_state(child, spec) for child in children))


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


def rendered(style):
    if style['fill_pattern'] in ('solid', 'none'):
        style['fill_background'] = None
    return style


def facet(style, flag):
    return tuple(style[key] for key in FACETS[flag])


def restyle(style, patch):
    rendered(style)
    for flag, keys in FACETS.items():
        if set(keys) & set(patch):
            if facet(style, flag) != facet(style['base'], flag):
                style['applied'][flag] = True
            else:
                style['applied'].pop(flag, None)


def indexed(items, index, kind):
    index = int(index)
    require(0 <= index < len(items), f'missing {kind} {index}')
    return items[index]


def styles(archive, real=False):
    part = related(archive, 'xl/workbook.xml', 'styles')
    require(part is not None, 'workbook styles are missing')
    root = xml(archive, part)
    formats = dict(BUILTIN_FORMATS)
    formats.update({int(n.attrib['numFmtId']): n.attrib['formatCode'] for n in root.findall('s:numFmts/s:numFmt', NS)})
    fonts = root.findall('s:fonts/s:font', NS)
    fills = root.findall('s:fills/s:fill', NS)
    borders = root.findall('s:borders/s:border', NS)

    def resolve(xf, modeled):
        font = copy.deepcopy(fonts[int(xf.get('fontId', '0'))])
        bold, italic = boolean(font.find('s:b', NS)), boolean(font.find('s:i', NS))
        for tag in ('b', 'i'):
            nodes = font.findall(f's:{tag}', NS)
            require(len(nodes) <= 1, f'duplicate font child {tag}')
            for node in nodes:
                font.remove(node)
        fill = copy.deepcopy(fills[int(xf.get('fillId', '0'))])
        pattern = fill.find('s:patternFill', NS)
        color = pattern.find('s:fgColor', NS) if pattern is not None else None
        background = pattern.find('s:bgColor', NS) if pattern is not None else None
        kind = pattern.get('patternType', 'none') if pattern is not None else 'none'
        if pattern is not None:
            pattern.attrib.pop('patternType', None)
            for node in (color, background):
                if node is not None:
                    pattern.remove(node)
        fmt = int(xf.get('numFmtId', '0'))
        require(fmt in formats or (real and 0 <= fmt < 164), f'unknown number format {fmt}')
        style = rendered({'number_format': formats.get(fmt, f'builtin:{fmt}'), 'bold': bold, 'italic': italic, 'fill_pattern': kind,
                          'fill': color.get('rgb', '').upper() if color is not None else None,
                          'fill_other': tuple(sorted((k, v) for k, v in color.attrib.items() if k != 'rgb')) if color is not None else (),
                          'fill_background': canonical(background), 'fill_rest': canonical(fill),
                          'font': canonical(font), 'border': canonical(borders[int(xf.get('borderId', '0'))]),
                          'alignment': canonical(xf.find('s:alignment', NS)),
                          'protection': canonical(xf.find('s:protection', NS))})
        attributes = {key: value for key, value in xf.attrib.items() if key not in POOLS and key not in modeled}
        for key in ('quotePrefix', 'pivotButton'):
            attributes[key] = attributes.get(key, '0') in ('1', 'true')
        style['xf'] = {'attributes': attributes,
                       'children': tuple(canonical(child) for child in xf if child.tag not in XF_FACET_CHILDREN)}
        return style

    named = [resolve(xf, ()) for xf in root.findall('s:cellStyleXfs/s:xf', NS)]
    result = []
    for xf in root.findall('s:cellXfs/s:xf', NS):
        style = resolve(xf, (*FACETS, 'xfId'))
        style['base'] = indexed(named, xf.get('xfId', '0'), 'cell style format')
        style['applied'] = {flag: xf.get(flag, '1') in ('1', 'true') for flag in FACETS
                            if facet(style, flag) != facet(style['base'], flag)}
        if real:
            for flag, keys in FACETS.items():
                if xf.get(flag) in ('0', 'false'):
                    for key in keys:
                        style[key] = copy.deepcopy(style['base'][key])
            style['applied'] = {flag: True for flag in FACETS if facet(style, flag) != facet(style['base'], flag)}
        result.append(style)
    require(bool(result), 'no cell styles')
    named_styles = {}
    for item in root.findall('s:cellStyles/s:cellStyle', NS):
        name = item.get('name')
        require(name not in named_styles, f'duplicate cell style {name}')
        named_styles[name] = {'attributes': {key: value for key, value in item.attrib.items() if key != 'xfId'},
                              'children': tuple(canonical(child) for child in item),
                              'style': indexed(named, item.get('xfId', '0'), 'cell style format')}
    return result, named_styles


def workbook(archive, real=False):
    cell_styles, named_styles = styles(archive, real)
    strings_part = related(archive, 'xl/workbook.xml', 'sharedStrings')
    string_root = xml(archive, strings_part) if strings_part else None
    string_nodes = string_root.findall('s:si', NS) if string_root is not None else []
    strings = [(''.join(n.text or '' for n in item.iter(f'{{{NS["s"]}}}t')) if real else ''.join(item.itertext())) for item in string_nodes]
    links = relations(archive, 'xl/workbook.xml')
    root = xml(archive, 'xl/workbook.xml')
    modeled = {'xl/workbook.xml', related(archive, 'xl/workbook.xml', 'styles')}
    style_state = xml(archive, related(archive, 'xl/workbook.xml', 'styles'))
    for child in list(style_state):
        if child.tag in {f'{{{NS["s"]}}}{kind}' for kind in ('numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles')}:
            style_state.remove(child)
    if strings_part:
        modeled.add(strings_part)
    result, sheet_views = [], []
    for entry in root.findall('s:sheets/s:sheet', NS):
        link = links[entry.attrib[f'{{{NS["r"]}}}id']]
        require(link['type'] == 'worksheet' and not link['external'], 'sheet must reference a worksheet')
        sheet = xml(archive, link['target'])
        modeled.add(link['target'])
        cells = {}
        row_style = {row.get('r'): row.get('s', '0') for row in sheet.findall('s:sheetData/s:row', NS) if row.get('customFormat') in ('1', 'true')} if real else {}
        column_styles = [(int(col.get('min')), int(col.get('max')), col.get('style', '0')) for col in sheet.findall('s:cols/s:col', NS)] if real else []
        shared = {n.find('s:f', NS).get('si'): (n.get('r'), n.find('s:f', NS))
                  for n in sheet.findall('s:sheetData/s:row/s:c', NS)
                  if real and n.find('s:f', NS) is not None and n.find('s:f', NS).get('t') == 'shared' and n.find('s:f', NS).text}
        for node in sheet.findall('s:sheetData/s:row/s:c', NS):
            address = node.attrib['r']
            require(address not in cells, f'duplicate cell {address}')
            kind, value = node.get('t', 'n'), node.findtext('s:v', namespaces=NS)
            formula = node.find('s:f', NS)
            require(real or formula is None or formula.get('t', 'normal') == 'normal', 'shared or array formula cannot be checked')
            formula_text = formula.text if formula is not None else None
            if real and formula is not None and formula.get('t') == 'shared':
                require(formula.get('si') in shared, 'missing shared formula master')
                origin, master = shared[formula.get('si')]
                formula_text = translate_formula(master.text, origin, address)
            if kind == 'inlineStr':
                value = ''.join(n.text or '' for n in node.findall('s:is//s:t', NS))
            elif kind == 's':
                value = strings[int(value)]
            elif kind == 'b':
                value = value == '1'
            elif kind == 'n' and value is not None:
                require(value.strip(), f'{entry.attrib["name"]}!{address}: empty cached value')
                value = Decimal(value)
                require(value.is_finite(), f'non-finite value in {address}')
                if real:
                    value = Decimal(str(float(value)))
                    require(value.is_finite(), f'non-finite double in {address}')
            elif kind == 'e':
                value = {'error': value}
            inherited = row_style.get(re.search(r'[0-9]+$', address).group())
            if real and inherited is None:
                column = column_number(re.match('[A-Z]+', address).group())
                inherited = next((style for low, high, style in column_styles if low <= column <= high), '0')
            style = copy.deepcopy(indexed(cell_styles, node.get('s', inherited or '0'), 'cell format'))
            record = {'value': value, 'formula': formula_text, 'style': style}
            if real:
                record['attributes'] = {k: v for k, v in node.attrib.items() if k not in ('r', 's', 't')}
                record['formula_properties'] = {k: v for k, v in formula.attrib.items() if k not in ('si', 'ref', 't') or (k in ('t', 'ref') and formula.get('t') != 'shared')} if formula is not None else {}
                if record['formula_properties'].get('t') == 'normal':
                    record['formula_properties'].pop('t')
                record['other'] = tuple(canonical(n) for n in node if n.tag not in {f'{{{NS["s"]}}}{t}' for t in ('f', 'v', 'is')})
                text_node = node.find('s:is', NS) if kind == 'inlineStr' else (string_nodes[int(node.findtext('s:v', namespaces=NS))] if kind == 's' else None)
                record['rich_text'] = rich_cell(text_node) if text_node is not None else None
            if value is not None or formula is not None or style != cell_styles[0] or (real and (record['attributes'] or record['other'] or record['rich_text'])):
                cells[address] = record
        views = sheet.findall('s:sheetViews/s:sheetView', NS)
        sheet_views.append([(int(view.get('workbookViewId', '0')), view.get('tabSelected') in ('1', 'true')) for view in views])
        for view in views:
            view.attrib.pop('tabSelected', None)
        metadata = []
        if real:
            data = sheet.find('s:sheetData', NS)
            metadata.append(('sheetData attributes', dict(data.attrib) if data is not None else {}))
            metadata.append(('row extensions', {row.get('r'): tuple(canonical(child) for child in row if child.tag != f'{{{NS["s"]}}}c') for row in data if any(child.tag != f'{{{NS["s"]}}}c' for child in row)} if data is not None else {}))
        for node in sheet:
            if node.tag in (f'{{{NS["s"]}}}sheetData', f'{{{NS["s"]}}}dimension'):
                continue
            if node.tag == f'{{{NS["s"]}}}cols':
                node = copy.deepcopy(node)
                formats = [indexed(cell_styles, column.attrib.pop('style', '0'), 'column format') for column in node]
                metadata.append((canonical(node), formats))
            else:
                metadata.append(canonical(node))
        rows = {}
        for node in sheet.findall('s:sheetData/s:row', NS):
            properties = {k: v for k, v in node.attrib.items() if k not in ('r', 'spans', 's')}
            style = indexed(cell_styles, node.get('s', '0'), 'row format')
            if style != cell_styles[0]:
                properties['s'] = style
            if properties:
                rows[node.attrib['r']] = properties
        result.append({'name': entry.attrib['name'], 'state': entry.get('state', 'visible'),
                       'properties': {key: value for key, value in entry.attrib.items()
                                      if key not in ('name', 'state', f'{{{NS["r"]}}}id')},
                       'relationship_properties': relationship_properties(archive, 'xl/workbook.xml', entry.attrib[f'{{{NS["r"]}}}id']),
                       'content_type': content_type(archive, link['target']),
                       'sheet_attributes': dict(sheet.attrib),
                       'cells': cells, 'metadata': metadata, 'row_properties': rows})
    require(len({sheet['name'] for sheet in result}) == len(result), 'duplicate worksheet names')
    extra = {}
    if real and string_root is not None:
        string_root.attrib.pop('count', None)
        string_root.attrib.pop('uniqueCount', None)
        for item in list(string_root):
            if item.tag == f'{{{NS["s"]}}}si':
                string_root.remove(item)
        extra['strings_state'] = canonical(string_root)
    elif real:
        extra['strings_state'] = canonical(ET.Element(f'{{{NS["s"]}}}sst'))
    return {**extra, 'sheets': result, 'default_style': cell_styles[0], 'workbook_state': workbook_state(archive, root),
            'selection': selection_state(root, [sheet['name'] for sheet in result], sheet_views),
            'named_styles': named_styles, 'style_state': canonical(style_state),
            'content_types': {'workbook': content_type(archive, 'xl/workbook.xml'),
                              'styles': content_type(archive, related(archive, 'xl/workbook.xml', 'styles'))},
            'package_state': package_state(archive, modeled, {'xl/workbook.xml': ('worksheet', 'sharedStrings')})}


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
    expected['workbook_state'] = expected_workbook_state(source['workbook_state'], spec)
    renames = spec.get('rename', {})
    expected['selection'] = {view: sorted(renames.get(name, name) for name in names)
                             for view, names in source['selection'].items()}
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
                    restyle(record['style'], value)
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
    for child in node:
        kind = child.tag.removeprefix(f'{{{NS["a"]}}}')
        style = canonical(child.find('a:rPr', NS))
        text = child.findtext('a:t', default='', namespaces=NS)
        if kind == 'r':
            if runs and runs[-1]['kind'] == 'r' and runs[-1]['style'] == style:
                runs[-1]['text'] += text
            else:
                runs.append({'kind': kind, 'text': text, 'style': style, 'field': None})
        elif kind == 'br':
            runs.append({'kind': kind, 'text': '\n', 'style': style, 'field': None})
        elif kind == 'fld':
            runs.append({'kind': kind, 'text': text, 'style': style,
                         'field': {'attributes': dict(child.attrib), 'props': canonical(child.find('a:pPr', NS))}})
        else:
            require(kind in ('pPr', 'endParaRPr'), f'unexpected paragraph child {child.tag}')
    return {'text': ''.join(run['text'] for run in runs), 'bullet': bool(bullet), 'props': canonical(props),
            'runs': runs, 'end': canonical(node.find('a:endParaRPr', NS))}


def notes_part(archive, part):
    target = related(archive, part, 'notesSlide')
    if not target:
        return None
    root = resolved_xml(archive, target)
    tree = root.find('p:cSld/p:spTree', NS)
    require(tree is not None, 'notes slide has no shape tree')
    shapes = []
    for node in list(tree):
        body = node.find('p:txBody', NS)
        paragraphs = body.findall('a:p', NS) if body is not None else []
        for item in paragraphs:
            body.remove(item)
        placeholder = node.find('p:nvSpPr/p:nvPr/p:ph', NS)
        shapes.append({'element': canonical(node), 'paragraphs': [paragraph(item) for item in paragraphs],
                       'notes_body': node.tag == f'{{{NS["p"]}}}sp' and (placeholder is None or placeholder.get('type') == 'body')})
        tree.remove(node)
    return {'content_type': content_type(archive, target), 'relationships': relationship_state(archive, target, owner=part),
            'document': canonical(root), 'shapes': shapes}


def notes_text(notes):
    if notes is None:
        return []
    return [item['text'] for shape in notes['shapes'] if shape['notes_body'] for item in shape['paragraphs']]


def presentation(archive):
    root = xml(archive, 'ppt/presentation.xml')
    links = relations(archive, 'ppt/presentation.xml')
    modeled = {'ppt/presentation.xml'}
    dynamic_links = {'ppt/presentation.xml': ('slide',)}
    for name in archive.namelist():
        if name.endswith('.xml') and name.startswith(('ppt/slides/', 'ppt/notesSlides/')):
            modeled.add(name)
            modeled.add(relpath(name))
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
            element = resolved_xml(archive, part, node)
            body = element.find('p:txBody', NS)
            for item in body.findall('a:p', NS) if body is not None else []:
                body.remove(item)
            shapes.append({'name': identity.get('name', ''), 'rect': rect, 'properties': canonical(props),
                           'transform_properties': dict(transform.attrib) if transform is not None else {},
                           'placeholder': canonical(placeholder), 'body_properties': canonical(node.find('p:txBody/a:bodyPr', NS)),
                           'paragraphs': [paragraph(n, inherited) for n in node.findall('p:txBody/a:p', NS)],
                           'element': canonical(element)})
        notes = notes_part(archive, part)
        slides.append({'shapes': shapes, 'notes': notes_text(notes), 'notes_part': notes, 'layout': layout_part,
                       'identity': {key: value for key, value in entry.attrib.items() if key != f'{{{NS["r"]}}}id'},
                       'relationship_properties': relationship_properties(archive, 'ppt/presentation.xml', entry.attrib[f'{{{NS["r"]}}}id']),
                       'content_type': content_type(archive, part),
                       'slide_attributes': dict(slide.attrib),
                       'common_attributes': dict(slide.find('p:cSld', NS).attrib),
                       'relationships': relationship_state(archive, part, ('notesSlide',)),
                       'background': canonical(slide.find('p:cSld/p:bg', NS)),
                       'group_properties': [canonical(node) for node in tree if node.tag in (f'{{{NS["p"]}}}nvGrpSpPr', f'{{{NS["p"]}}}grpSpPr')],
                       'slide_properties': [canonical(node) for node in slide if node.tag != f'{{{NS["p"]}}}cSld']})
    state = resolved_xml(archive, 'ppt/presentation.xml', root)
    state.find('p:sldIdLst', NS)[:] = []
    return {'slides': slides, 'size': dict(root.find('p:sldSz', NS).attrib),
            'presentation_state': canonical(state), 'package_state': package_state(archive, modeled, dynamic_links),
            'content_type': content_type(archive, 'ppt/presentation.xml'),
            'assets': {name: canonical(xml(archive, name)) for name in archive.namelist()
                       if name.endswith('.xml') and name.startswith(('ppt/slideLayouts/', 'ppt/slideMasters/', 'ppt/notesMasters/', 'ppt/theme/'))}}


def styled_run(item):
    return len(item['runs']) == 1 and item['runs'][0]['kind'] == 'r'


def replace_paragraph_text(item, text):
    if text == item['text']:
        return
    require(styled_run(item), 'fixture replacement must have one styled run')
    item['text'] = text
    item['runs'][0]['text'] = text


def replace_notes(notes, lines):
    bodies = [shape for shape in notes['shapes'] if shape['notes_body']]
    require(len(bodies) == 1, 'fixture notes must have one body')
    paragraphs = bodies[0]['paragraphs']
    require(paragraphs and all(styled_run(item) for item in paragraphs)
            and len({(item['props'], item['runs'][0]['style'], item['end']) for item in paragraphs}) == 1,
            'fixture notes must have uniform paragraphs with one styled run')
    template = paragraphs[0]
    run = template['runs'][0]
    bodies[0]['paragraphs'] = [{'text': line, 'bullet': template['bullet'], 'props': template['props'], 'end': template['end'],
                                'runs': [dict(run, text=line)]}
                               for line in lines]
    return notes


def check_created_notes(notes, path):
    require(notes is not None, f'{path}: notes part is missing')
    compare(NOTES_TYPE, notes['content_type'], f'{path}.notes content type')
    require(any(dict(link)['Target'] == OWNER and dict(link)['Type'].rsplit('/', 1)[-1] == 'slide'
                for link in notes['relationships']), f'{path}: notes part does not link back to its slide')


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
            slide['notes_part'] = replace_notes(slide['notes_part'], slide['notes']) if slide['notes_part'] else CREATED
        slide['shapes'].extend({'requested': request} for request in spec.get('add_shapes', {}).get(str(i), []))
    if 'slide_order' in spec:
        expected['slides'] = [expected['slides'][item - 1] if isinstance(item, int) else {'requested': item} for item in spec['slide_order']]
    return expected


def compare(expected, actual, path='document'):
    if isinstance(expected, NewParagraphAttributes):
        expected.compare(actual, path)
    elif isinstance(expected, dict):
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
    compare(expected['presentation_state'], actual['presentation_state'], 'presentation state')
    compare(expected['package_state'], actual['package_state'], 'presentation package')
    compare(expected['content_type'], actual['content_type'], 'presentation content type')
    require(len(expected['slides']) == len(actual['slides']), 'slide count differs')
    for i, (before, after) in enumerate(zip(expected['slides'], actual['slides']), 1):
        path = f'slide {i}'
        if 'requested' in before:
            request = before['requested']
            notes = request['notes'].split('\n') if request.get('notes') else []
            compare(notes, after['notes'], path + '.notes')
            if notes:
                check_created_notes(after['notes_part'], path)
            require(len(request['shapes']) == len(after['shapes']), path + ': shape count differs')
            for j, (shape, found) in enumerate(zip(request['shapes'], after['shapes']), 1):
                check_requested_shape(shape, found, f'{path}.shape {j}')
        else:
            for key in ('identity', 'relationship_properties', 'content_type', 'slide_attributes',
                        'common_attributes', 'notes', 'layout', 'relationships',
                        'background', 'group_properties', 'slide_properties'):
                compare(before[key], after[key], f'{path}.{key}')
            if before['notes_part'] == CREATED:
                check_created_notes(after['notes_part'], path)
            else:
                compare(before['notes_part'], after['notes_part'], f'{path}.notes_part')
            require(len(before['shapes']) == len(after['shapes']), path + ': shape count differs')
            for j, (shape, found) in enumerate(zip(before['shapes'], after['shapes']), 1):
                if 'requested' in shape:
                    check_requested_shape(shape['requested'], found, f'{path}.shape {j}')
                else:
                    compare(shape, found, f'{path}.shape {j}')


def w(name):
    return f'{{{NS["w"]}}}{name}'


def local(tag):
    return tag.rsplit('}', 1)[-1]


def properties(node):
    return None if node is None or (not node.attrib and len(node) == 0) else canonical(node)


def run_model(node):
    content = []
    for child in node:
        if child.tag == w('t'):
            text = child.text or ''
            space = child.get('{http://www.w3.org/XML/1998/namespace}space')
            if text and (text[0].isspace() or text[-1].isspace()) and space != 'preserve':
                content.append(('xml:space', None))
            content.append(('text', text))
        elif child.tag != w('rPr'):
            content.append(canonical(child))
    return {'kind': 'r', 'attributes': dict(node.attrib), 'properties': properties(node.find('w:rPr', NS)), 'content': content}


def joined(content):
    result = []
    for token in content:
        if token[0] == 'text' and result and result[-1][0] == 'text':
            result[-1] = ('text', result[-1][1] + token[1])
        elif token != ('text', ''):
            result.append(token)
    return result


def merged_runs(items):
    result = []
    for item in items:
        previous = result[-1] if result else None
        if previous and previous['kind'] == item['kind'] == 'r' and \
                (previous['attributes'], previous['properties']) == (item['attributes'], item['properties']):
            previous['content'] += item['content']
        else:
            result.append(item)
    for item in result:
        if item['kind'] == 'r':
            item['content'] = joined(item['content'])
    return result


def paragraph_text(item):
    return ''.join(token[1] for run in item['content'] if run['kind'] == 'r' for token in run['content'] if token[0] == 'text')


def docx_paragraph(node, default_style):
    props = copy.deepcopy(node.find('w:pPr', NS))
    style = props.find('w:pStyle', NS) if props is not None else None
    if style is not None:
        props.remove(style)
    item = {'kind': 'p', 'attributes': dict(node.attrib), 'style': style.get(w('val')) if style is not None else default_style,
            'properties': properties(props),
            'content': merged_runs([run_model(child) if child.tag == w('r') else {'kind': local(child.tag), 'element': canonical(child)}
                                    for child in node if child.tag != w('pPr')])}
    item['text'] = paragraph_text(item)
    return item


def docx_blocks(children, default_style):
    blocks = []
    for child in children:
        if child.tag == w('p'):
            blocks.append(docx_paragraph(child, default_style))
        elif child.tag == w('tbl'):
            rows = [{'attributes': dict(row.attrib), 'properties': [canonical(item) for item in row if item.tag != w('tc')],
                     'cells': [{'attributes': dict(cell.attrib), 'properties': properties(cell.find('w:tcPr', NS)),
                                'blocks': docx_blocks([item for item in cell if item.tag != w('tcPr')], default_style)}
                               for cell in row.findall('w:tc', NS)]}
                    for row in child.findall('w:tr', NS)]
            blocks.append({'kind': 'tbl', 'attributes': dict(child.attrib),
                           'properties': [canonical(item) for item in child if item.tag != w('tr')], 'rows': rows})
        else:
            blocks.append({'kind': local(child.tag), 'element': canonical(child)})
    return blocks


def tracked_changes(archive):
    counts = {}
    for name in archive.namelist():
        if name.endswith('.xml'):
            for node in xml(archive, name).iter():
                if node.tag in TRACKED:
                    counts[local(node.tag)] = counts.get(local(node.tag), 0) + 1
    return counts


def document(archive):
    part = related(archive, '', 'officeDocument')
    require(part is not None, 'main document part is missing')
    styles_part = related(archive, part, 'styles')
    defaults = [style.get(w('styleId')) for style in xml(archive, styles_part).findall('w:style', NS)
                if style.get(w('type')) == 'paragraph' and style.get(w('default')) in ('1', 'true', 'on')] if styles_part else []
    require(len(defaults) <= 1, 'multiple default paragraph styles')
    default_style = defaults[0] if defaults else None
    root = resolved_xml(archive, part)
    body = root.find('w:body', NS)
    require(body is not None, 'document has no body')
    blocks = docx_blocks(body, default_style)
    root.remove(body)
    state = package_state(archive, {part}, {})
    core = related(archive, '', 'core-properties')
    if core in state:
        saved = resolved_xml(archive, core)
        for node in saved.iter(SAVE_TIME):
            node.text = 'updated on save'
        state[core] = canonical(saved)
    return {'content_type': content_type(archive, part), 'document': canonical(root), 'body_attributes': dict(body.attrib),
            'body': blocks, 'default_style': default_style, 'tracked_changes': tracked_changes(archive), 'package_state': state}


def docx_paragraphs(blocks):
    for block in blocks:
        if block['kind'] == 'p':
            yield block
        elif block['kind'] == 'tbl':
            for row in block['rows']:
                for cell in row['cells']:
                    yield from docx_paragraphs(cell['blocks'])


class NewParagraphAttributes(dict):
    def __init__(self, attributes, input_ids):
        super().__init__((key, value) for key, value in attributes.items() if key not in PARAGRAPH_IDS)
        self.input_ids = input_ids

    def compare(self, actual, path):
        require(isinstance(actual, dict), f'{path}: fields differ')
        for key in PARAGRAPH_IDS:
            if key in actual:
                value = actual[key]
                require(re.fullmatch(r'[0-9A-Fa-f]{8}', value) and 0 < int(value, 16) < 0x80000000,
                        f'{path}: invalid {local(key)} {value!r}')
                require(value.upper() not in self.input_ids, f'{path}: {local(key)} reuses an input paragraph id')
        compare(dict(self), {key: value for key, value in actual.items() if key not in PARAGRAPH_IDS}, path)


def docx_unique(blocks, text):
    matches = [item for item in docx_paragraphs(blocks) if item['text'] == text]
    require(len(matches) == 1, f'fixture must have exactly one paragraph with text {text!r}')
    return matches[0]


def top_level(body, text):
    item = docx_unique(body, text)
    require(any(block is item for block in body), f'fixture paragraph {text!r} must be in the body')
    return item


def set_docx_text(item, text):
    require(len(item['content']) == 1 and item['content'][0]['kind'] == 'r' and item['content'][0]['content'] == [('text', item['text'])],
            'fixture paragraph must have one plain text run')
    item['content'][0]['content'] = [('text', text)]
    item['text'] = text


def with_bold(props):
    children = [child for child in (props[3] if props else ()) if child[0] != w('b')]
    children.insert(sum(child[0] in (w('rStyle'), w('rFonts')) for child in children), (w('b'), (), None, ()))
    return (w('rPr'), props[1] if props else (), None, tuple(children))


def bold_phrase(item, phrase):
    for index, run in enumerate(item['content']):
        if run['kind'] == 'r' and len(run['content']) == 1 and run['content'][0][0] == 'text' and phrase in run['content'][0][1]:
            before, after = run['content'][0][1].split(phrase, 1)
            item['content'][index:index + 1] = [dict(run, properties=props, content=[('text', text)]) for text, props in
                                                ((before, run['properties']), (phrase, with_bold(run['properties'])), (after, run['properties'])) if text]
            return
    require(False, f'fixture phrase {phrase!r} must be inside one text run')


def docx_cell(body, table, row, column):
    tables = [block for block in body if block['kind'] == 'tbl']
    require(0 < table <= len(tables), f'fixture table {table} is missing')
    return tables[table - 1]['rows'][row - 1]['cells'][column - 1]


def cell_paragraph(cell):
    require(len(cell['blocks']) == 1 and cell['blocks'][0]['kind'] == 'p', 'fixture cell must hold one paragraph')
    return cell['blocks'][0]


def at(blocks, item):
    return next(i for i, block in enumerate(blocks) if block is item)


def expected_document(source, spec):
    expected = copy.deepcopy(source)
    body = expected['body']
    input_ids = {value.upper() for item in docx_paragraphs(source['body'])
                 for key, value in item['attributes'].items() if key in PARAGRAPH_IDS}
    styled = [(top_level(body, text), style) for text, style in spec.get('styles', {}).items()]
    bolded = [(docx_unique(body, item['paragraph']), item['phrase']) for item in spec.get('bold', [])]
    texts = [(docx_unique(body, old), new) for old, new in spec.get('text', {}).items()]
    texts += [(cell_paragraph(docx_cell(body, item['table'], item['row'], item['column'])), item['text']) for item in spec.get('cells', [])]
    deleted = [top_level(body, text) for text in spec.get('delete_paragraphs', [])]
    anchors = [(top_level(body, text), lines) for text, lines in spec.get('insert_paragraphs', {}).items()]
    if 'replace' in spec:
        old, new = spec['replace']['old'], spec['replace']['new']
        for item in docx_paragraphs(body):
            if old in item['text']:
                wanted = item['text'].replace(old, new)
                for run in item['content']:
                    if run['kind'] == 'r':
                        run['content'] = [('text', token[1].replace(old, new)) if token[0] == 'text' else token for token in run['content']]
                item['text'] = paragraph_text(item)
                require(item['text'] == wanted, f'fixture occurrence of {old!r} must be inside one text run')
    for item, text in texts:
        set_docx_text(item, text)
    for item, phrase in bolded:
        bold_phrase(item, phrase)
    for item, style in styled:
        item['style'] = style
    tables = [block for block in body if block['kind'] == 'tbl']
    for change in spec.get('insert_rows', []):
        rows = tables[change['table'] - 1]['rows']
        row = copy.deepcopy(rows[change['like'] - 1])
        for cell, text in zip(row['cells'], change['cells'], strict=True):
            set_docx_text(cell_paragraph(cell), text)
            for item in docx_paragraphs(cell['blocks']):
                item['attributes'] = NewParagraphAttributes(item['attributes'], input_ids)
        rows.insert(change['before'] - 1, row)
    for change in spec.get('delete_rows', []):
        del tables[change['table'] - 1]['rows'][change['row'] - 1]
    for item in deleted:
        del body[at(body, item)]
    for item, lines in anchors:
        position = at(body, item) + 1
        body[position:position] = [{'kind': 'p', 'attributes': NewParagraphAttributes({}, input_ids), 'style': expected['default_style'], 'properties': None, 'text': line,
                                    'content': [{'kind': 'r', 'attributes': {}, 'properties': None, 'content': [('text', line)]}]}
                                   for line in lines]
    return expected


def check_document(expected, actual):
    for kind, found in actual['tracked_changes'].items():
        require(found <= expected['tracked_changes'].get(kind, 0), f'tracked changes are not allowed: w:{kind} added')
    compare(expected, actual)
    for key in PARAGRAPH_IDS:
        seen = set()
        for item in docx_paragraphs(actual['body']):
            value = item['attributes'].get(key)
            if value is not None:
                require(value.upper() not in seen, f'document.body: duplicate {local(key)} {value!r}')
                seen.add(value.upper())


def translate_formula(formula, origin, destination):
    def coordinates(address):
        col, row = re.fullmatch(r'([A-Z]+)([0-9]+)', address).groups()
        return column_number(col), int(row)
    before, after = coordinates(origin), coordinates(destination)
    def shift(match):
        absolute_col, col, absolute_row, row = match.groups()
        col = column_number(col) + (0 if absolute_col else after[0] - before[0])
        row = int(row) + (0 if absolute_row else after[1] - before[1])
        require(col > 0 and row > 0, 'invalid translated shared formula reference')
        return absolute_col + column_letters(col) + absolute_row + str(row)
    pieces = re.split(r'("(?:[^"]|"")*")', formula)
    for i in range(0, len(pieces), 2):
        pieces[i] = re.sub(r'(?<![\w.])(?!(?:[A-Z]+[0-9]+)\s*\()(\$?)([A-Z]{1,3})(\$?)([0-9]+)(?![\w.!])', shift, pieces[i])
    return ''.join(pieces)


def rich_cell(node):
    if node is None:
        return None
    runs = []
    for run in node.findall('s:r', NS):
        style = properties(run.find('s:rPr', NS))
        text = run.findtext('s:t', default='', namespaces=NS)
        extra = (tuple(sorted(run.attrib.items())), tuple(canonical(child) for child in run if child.tag not in {f'{{{NS["s"]}}}rPr', f'{{{NS["s"]}}}t'}))
        if runs and runs[-1][:2] == (style, extra):
            runs[-1] = (style, extra, runs[-1][2] + text)
        else:
            runs.append((style, extra, text))
    rest = (tuple(sorted(node.attrib.items())), tuple(canonical(child) for child in node if child.tag not in {f'{{{NS["s"]}}}r', f'{{{NS["s"]}}}t'}))
    if not any(style is not None or extra != ((), ()) for style, extra, _ in runs) and rest == ((), ()):
        return None
    return (tuple(runs), rest)


class RealPackage:
    def __init__(self, archive):
        self.parts = {name: archive.read(name) for name in archive.namelist() if not name.endswith('/')}
        self.roots = {}
        self.namespaces = {}

    def namelist(self):
        return list(self.parts)

    def read(self, name):
        return ET.tostring(self.roots[name]) if name in self.roots else self.parts[name]

    def root(self, name):
        if name not in self.roots:
            import io
            self.namespaces[name] = dict(item for _, item in ET.iterparse(io.BytesIO(self.parts[name]), events=('start-ns',)))
            self.roots[name] = ET.fromstring(self.parts[name])
        return self.roots[name]


def paragraph_nodes(root, prefix):
    return list(root.iter(f'{{{NS[prefix]}}}p'))


def node_text(node, prefix):
    return ''.join(n.text or '' for n in node.iter(f'{{{NS[prefix]}}}t'))


def replace_xml_text(paragraph, old, new, prefix):
    nodes = list(paragraph.iter(f'{{{NS[prefix]}}}t'))
    text = ''.join(n.text or '' for n in nodes)
    starts = [m.start() for m in re.finditer(re.escape(old), text)]
    for start in reversed(starts):
        offset = 0
        end = start + len(old)
        for node in nodes:
            value = node.text or ''
            left, right = max(0, start - offset), min(len(value), end - offset)
            if left < right:
                node.text = value[:left] + (new if offset <= start < offset + len(value) else '') + value[right:]
                if prefix == 'w' and node.text and (node.text[0] in ' \t\r\n' or node.text[-1] in ' \t\r\n'):
                    node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            offset += len(value)
    return len(starts)


def accept_revisions(root):
    discard = {w(name) for name in ('del', 'moveFrom', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd',
                                   'rPrChange', 'pPrChange', 'sectPrChange', 'tblPrChange', 'tblPrExChange', 'tblGridChange',
                                   'trPrChange', 'tcPrChange', 'cellDel', 'numberingChange')}
    unwrap = {w(name) for name in ('ins', 'moveTo')}
    for parent in root.iter():
        for child in list(parent):
            if child.tag in discard:
                parent.remove(child)
            elif child.tag in unwrap:
                index = list(parent).index(child)
                parent.remove(child)
                parent[index:index] = list(child)
            elif child.tag in {w('cellIns'), w('cellMerge')}:
                raise ValueError('unsupported table revision in fixture')


def real_expected_package(package, spec, fmt):
    prefix = 'w' if fmt == 'docx' else 'a'
    if spec.get('accept_changes'):
        require(bool(tracked_changes(package)), 'fixture contains no Word revisions to accept')
        for name in package.namelist():
            if name.startswith('word/') and name.endswith('.xml'):
                accept_revisions(package.root(name))
    if 'insert_paragraphs' in spec:
        change = spec['insert_paragraphs']
        body = package.root('word/document.xml').find('w:body', NS)
        headings = [p for p in body if p.tag == w('p') and node_text(p, 'w') == change['heading']]
        require(len(headings) == 1, 'insertion heading must be unique')
        index = list(body).index(headings[0]) + 1
        template = body[index]
        require(template.tag == w('p'), 'insertion requires a following body paragraph')
        for offset, text in enumerate(change['texts']):
            paragraph = ET.Element(w('p'))
            props = template.find('w:pPr', NS)
            if props is not None:
                paragraph.append(copy.deepcopy(props))
            run = ET.SubElement(paragraph, w('r'))
            props = template.find('w:r/w:rPr', NS)
            if props is not None:
                run.append(copy.deepcopy(props))
            ET.SubElement(run, w('t')).text = text
            body.insert(index + offset, paragraph)
    if 'delete_paragraph' in spec:
        body = package.root('word/document.xml').find('w:body', NS)
        matches = [p for p in body if p.tag == w('p') and node_text(p, 'w') == spec['delete_paragraph']]
        require(len(matches) == 1, 'deleted paragraph must be unique')
        body.remove(matches[0])
    if 'format_phrase' in spec:
        change = spec['format_phrase']
        matches = [p for p in package.root('word/document.xml').iter(w('p')) if node_text(p, 'w') == change['paragraph']]
        require(len(matches) == 1, 'formatted paragraph must be unique')
        paragraph = matches[0]
        require(change['paragraph'].count(change['phrase']) == 1, 'formatted phrase must be unique')
        start = change['paragraph'].index(change['phrase'])
        end, position = start + len(change['phrase']), 0
        for run in list(paragraph):
            if run.tag != w('r'):
                continue
            text = node_text(run, 'w')
            low, high = max(0, start - position), min(len(text), end - position)
            position += len(text)
            if low >= high:
                continue
            require(len(run.findall('w:t', NS)) == 1 and all(n.tag in {w('rPr'), w('t')} for n in run), 'phrase run must contain only text')
            index = list(paragraph).index(run)
            paragraph.remove(run)
            pieces = [(text[:low], False), (text[low:high], True), (text[high:], False)]
            for value, formatted in pieces:
                if not value:
                    continue
                new = copy.deepcopy(run)
                node = new.find('w:t', NS)
                node.text = value
                if value[0] in ' \t\r\n' or value[-1] in ' \t\r\n':
                    node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                if formatted:
                    props = new.find('w:rPr', NS)
                    if props is None:
                        props = ET.Element(w('rPr'))
                        new.insert(0, props)
                    prop = props.find('w:' + change['property'], NS)
                    if prop is None:
                        prop = ET.SubElement(props, w(change['property']))
                    prop.set(w('val'), '1')
                paragraph.insert(index, new)
                index += 1
    for edit in spec.get('edits', []):
        matches = 0
        root = package.root(edit['part'])
        parents = {child: parent for parent in root.iter() for child in parent}
        for paragraph in paragraph_nodes(root, prefix):
            if any(ancestor.tag == f'{{{NS[prefix]}}}p' for ancestor in ancestors(paragraph, parents)):
                continue
            if 'paragraph' not in edit or node_text(paragraph, prefix) == edit['paragraph']:
                matches += replace_xml_text(paragraph, edit['old'], edit['new'], prefix)
        require(matches > 0, f'fixture edit does not match: {edit}')
    if 'slide_replace' in spec:
        change = spec['slide_replace']
        for name in package.namelist():
            if re.fullmatch(r'ppt/slides/slide[0-9]+.xml', name):
                for paragraph in paragraph_nodes(package.root(name), 'a'):
                    replace_xml_text(paragraph, change['old'], change['new'], 'a')
    for change in spec.get('table_cells', []):
        root = package.root(change['part'])
        table = root.findall('.//a:tbl', NS)[change['table'] - 1]
        cell = table.findall('a:tr', NS)[change['row'] - 1].findall('a:tc', NS)[change['column'] - 1]
        paragraphs = cell.findall('a:txBody/a:p', NS)
        require(len(paragraphs) == 1, 'table target must have one paragraph')
        replace_xml_text(paragraphs[0], node_text(paragraphs[0], 'a'), change['text'], 'a')
    for change in spec.get('delete_rows', []):
        root = package.root(change['part'])
        parents = {child: parent for parent in root.iter() for child in parent}
        rows = [row for row in root.iter(w('tr')) if change['contains'] in node_text(row, 'w')]
        require(len(rows) == 1, 'row deletion must identify exactly one row')
        parents[rows[0]].remove(rows[0])
    for text, style in spec.get('styles', {}).items():
        matches = [p for p in package.root('word/document.xml').iter(w('p')) if node_text(p, 'w') == text]
        require(len(matches) == 1, 'style target must be unique')
        props = matches[0].find('w:pPr', NS)
        if props is None:
            props = ET.Element(w('pPr'))
            matches[0].insert(0, props)
        item = props.find('w:pStyle', NS)
        if item is None:
            item = ET.Element(w('pStyle'))
            props.insert(0, item)
        item.set(w('val'), style)
    for slide, text in spec.get('notes', {}).items():
        part = slide_parts(package)[int(slide) - 1]
        target = related(package, part, 'notesSlide')
        require(target is not None, 'notes task requires an existing notes part')
        root = package.root(target)
        bodies = [shape.find('p:txBody', NS) for shape in root.findall('p:cSld/p:spTree/p:sp', NS)
                  if shape.find('p:nvSpPr/p:nvPr/p:ph', NS) is not None and shape.find('p:nvSpPr/p:nvPr/p:ph', NS).get('type') == 'body']
        require(len(bodies) == 1, 'notes task requires one body placeholder')
        body = bodies[0]
        paragraphs = body.findall('a:p', NS)
        require(bool(paragraphs), 'notes task requires a source paragraph')
        template = copy.deepcopy(paragraphs[0])
        for p in paragraphs:
            body.remove(p)
        for child in list(template):
            if child.tag not in {f'{{{NS["a"]}}}pPr', f'{{{NS["a"]}}}endParaRPr'}:
                template.remove(child)
        source_run = paragraphs[0].find('a:r', NS)
        run = ET.Element(f'{{{NS["a"]}}}r')
        if source_run is not None and source_run.find('a:rPr', NS) is not None:
            run.append(copy.deepcopy(source_run.find('a:rPr', NS)))
        ET.SubElement(run, f'{{{NS["a"]}}}t').text = text
        template.insert(1 if template.find('a:pPr', NS) is not None else 0, run)
        body.append(template)
    if spec.get('swap_last_slides'):
        listing = package.root('ppt/presentation.xml').find('p:sldIdLst', NS)
        listing[-2], listing[-1] = listing[-1], listing[-2]
    if 'move_slide' in spec:
        change = spec['move_slide']
        listing = package.root('ppt/presentation.xml').find('p:sldIdLst', NS)
        slide = listing[change['from'] - 1]
        listing.remove(slide)
        listing.insert(change['to'] - 1, slide)
    if 'delete_slide' in spec:
        listing = package.root('ppt/presentation.xml').find('p:sldIdLst', NS)
        slide = listing[spec['delete_slide'] - 1]
        identity = slide.get(f'{{{NS["r"]}}}id')
        part = relations(package, 'ppt/presentation.xml')[identity]['target']
        notes = related(package, part, 'notesSlide')
        listing.remove(slide)
        links = package.root(relpath('ppt/presentation.xml'))
        links.remove(next(n for n in links if n.get('Id') == identity))
        removed = {part, relpath(part)} | ({notes, relpath(notes)} if notes else set())
        types = package.root('[Content_Types].xml')
        for node in list(types):
            if node.get('PartName', '').lstrip('/') in removed:
                types.remove(node)
        for name in removed:
            package.parts.pop(name, None)
            package.roots.pop(name, None)


def ancestors(node, parents):
    while node in parents:
        node = parents[node]
        yield node


def slide_parts(archive):
    links = relations(archive, 'ppt/presentation.xml')
    return [links[n.get(f'{{{NS["r"]}}}id')]['target'] for n in xml(archive, 'ppt/presentation.xml').findall('p:sldIdLst/p:sldId', NS)]


class RealIdentity(NewParagraphAttributes):
    def __init__(self, attributes, input_ids, paragraph=True):
        dict.__init__(self, attributes)
        self.input_ids = input_ids
        self.paragraph = paragraph

    def compare(self, actual, path):
        expected = dict(self)
        ignorable = f'{{{NS["mc"]}}}Ignorable'
        actual = dict(actual)
        if not self.paragraph and NS['w14'] not in self.get(ignorable, '').split():
            tokens = set(actual.get(ignorable, '').split()) - {NS['w14']}
            if tokens:
                actual[ignorable] = ' '.join(sorted(tokens))
            else:
                actual.pop(ignorable, None)
        for key in PARAGRAPH_IDS if self.paragraph else ():
            if key not in self and key in actual:
                value = actual[key]
                require(re.fullmatch('[0-9A-Fa-f]{8}', value) and 0 < int(value, 16) < 0x80000000,
                        f'{path}: invalid fresh paragraph identity')
                require(value.upper() not in self.input_ids, f'{path}: new identity reuses an input id')
                expected[key] = value
        compare(expected, actual, path)


def real_metadata(package):
    for name in ('docProps/core.xml', 'docProps/app.xml'):
        if name not in package.parts:
            continue
        root = package.root(name)
        tags = {SAVE_TIME, '{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastModifiedBy'} if name.endswith('core.xml') else {f'{{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}}{key}' for key in ('Pages', 'Words', 'Characters', 'CharactersWithSpaces', 'Lines', 'Paragraphs', 'TotalTime', 'Slides', 'Notes', 'HiddenSlides', 'MMClips')}
        for child in list(root):
            if child.tag in tags:
                root.remove(child)


def effective_run_properties(node, styles_root, paragraph_style):
    definitions = {style.get(w('styleId')): style for style in styles_root.findall('w:style', NS)} if styles_root is not None else {}
    values = {}
    defaults = styles_root.find('w:docDefaults/w:rPrDefault/w:rPr', NS) if styles_root is not None else None
    toggles = {w(name) for name in ('b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'outline', 'shadow', 'emboss', 'imprint', 'vanish')}
    def apply(props, toggle=False):
        if props is not None:
            for child in props:
                if child.tag in toggles:
                    enabled = child.get(w('val'), '1') not in ('0', 'false', 'off')
                    if toggle and enabled:
                        enabled = child.tag not in values
                    if enabled:
                        values[child.tag] = (child.tag, (), None, ())
                    elif not toggle:
                        values.pop(child.tag, None)
                    elif child.get(w('val'), '1') not in ('0', 'false', 'off'):
                        values.pop(child.tag, None)
                elif child.tag != w('rStyle'):
                    values[child.tag] = canonical(child)
    def inherit(identity, seen):
        require(identity not in seen, 'cyclic Word style inheritance')
        style = definitions.get(identity)
        if style is not None:
            base = style.find('w:basedOn', NS)
            if base is not None:
                inherit(base.get(w('val')), seen | {identity})
            apply(style.find('w:rPr', NS), toggle=True)
    apply(defaults)
    inherit(paragraph_style, set())
    props = node.find('w:rPr', NS)
    char_style = props.find('w:rStyle', NS) if props is not None else None
    if char_style is not None:
        inherit(char_style.get(w('val')), set())
    apply(props)
    return tuple(sorted(values.items()))


def real_xml_state(package, name, expected=False, input_ids=frozenset()):
    root = copy.deepcopy(package.root(name))
    namespaces = package.namespaces[name]
    links = relations(package, name)
    styles_root = package.root('word/styles.xml') if 'word/styles.xml' in package.parts and name.startswith('word/') else None
    defaults = [s.get(w('styleId')) for s in styles_root.findall('w:style', NS) if s.get(w('type')) == 'paragraph' and s.get(w('default')) in ('1', 'true', 'on')] if styles_root is not None else []
    default_style = defaults[0] if defaults else None
    def model(node, paragraph_style=None):
        attrs = dict(node.attrib)
        for key, value in list(attrs.items()):
            if key.startswith(f'{{{NS["r"]}}}') and value:
                require(value in links, f'missing relationship ID in {name}: {value}')
                attrs[key] = json.dumps(links[value], sort_keys=True)
        if node.tag == f'{{{NS["mc"]}}}Choice' and 'Requires' in attrs:
            attrs['Requires'] = ' '.join(sorted(namespaces.get(value, value) for value in attrs['Requires'].split()))
        if f'{{{NS["mc"]}}}Ignorable' in attrs:
            key = f'{{{NS["mc"]}}}Ignorable'
            uris = {namespaces.get(value, value) for value in attrs[key].split()}
            if uris:
                attrs[key] = ' '.join(sorted(uris))
            else:
                del attrs[key]
        if expected and node is root and name.startswith('word/'):
            attrs = RealIdentity(attrs, input_ids, paragraph=False)
        if node.tag == w('p'):
            props = node.find('w:pPr', NS)
            style = props.find('w:pStyle', NS) if props is not None else None
            paragraph_style = style.get(w('val')) if style is not None else default_style
            if expected:
                attrs = RealIdentity(attrs, input_ids)
        tokens = []
        for child in node:
            if child.tag == w('proofErr'):
                continue
            if node.tag == w('p') and child.tag == w('pPr'):
                props = copy.deepcopy(child)
                item = props.find('w:pStyle', NS)
                if item is not None and item.get(w('val')) == default_style:
                    props.remove(item)
                if props.attrib or len(props):
                    tokens.append(model(props, paragraph_style))
            elif child.tag in (w('r'), f'{{{NS["a"]}}}r'):
                prefix = 'w' if child.tag == w('r') else 'a'
                run_attrs = {k: v for k, v in child.attrib.items() if not (prefix == 'w' and k in {w('rsidR'), w('rsidRPr'), w('rsidDel')})}
                props = effective_run_properties(child, styles_root, paragraph_style) if prefix == 'w' else properties(child.find('a:rPr', NS))
                content = []
                for item in child:
                    if item.tag == f'{{{NS[prefix]}}}t':
                        text = item.text or ''
                        if prefix == 'w' and text and (text[0] in ' \t\r\n' or text[-1] in ' \t\r\n'):
                            require(item.get('{http://www.w3.org/XML/1998/namespace}space') == 'preserve', 'significant Word whitespace lacks xml:space')
                        content.extend(('character', char, props, tuple(sorted(run_attrs.items()))) for char in text)
                    elif item.tag != f'{{{NS[prefix]}}}rPr':
                        content.append(('run-content', model(item, paragraph_style), props, tuple(sorted(run_attrs.items()))))
                tokens.extend(content or [('empty-run', props, tuple(sorted(run_attrs.items())))])
            else:
                tokens.append(model(child, paragraph_style))
        return {'tag': node.tag, 'attributes': attrs, 'text': node.text if node.text and (local(node.tag) in ('t', 'instrText', 'delText') or node.text.strip()) else None,
                'children': tokens}
    return model(root)


def real_ids(package):
    from collections import Counter
    return Counter((key, value.upper()) for name in package.namelist() if name.startswith('word/') and name.endswith('.xml')
                   for p in package.root(name).iter(w('p')) for key, value in p.attrib.items() if key in PARAGRAPH_IDS)


def real_package_state(package, expected=False, input_ids=frozenset()):
    real_metadata(package)
    state = {}
    for name in package.namelist():
        if name == 'xl/calcChain.xml':
            continue
        if name.endswith('.rels'):
            folder, filename = posixpath.split(name)
            part = posixpath.join(posixpath.dirname(folder), filename[:-5]) if name != '_rels/.rels' else ''
            records = relationship_state(package, part)
            if records:
                state[name] = records
        elif name == '[Content_Types].xml':
            state[name] = tuple(sorted(canonical(n) for n in package.root(name) if n.get('PartName') != '/xl/calcChain.xml'))
        elif name.endswith('.xml'):
            state[name] = real_xml_state(package, name, expected, input_ids)
        else:
            state[name] = hashlib.sha256(package.read(name)).hexdigest()
    return state


def real_slide_metadata(original, actual, spec):
    if not ({'move_slide', 'delete_slide', 'swap_last_slides'} & spec.keys()) or 'docProps/app.xml' not in original.parts:
        return
    root = original.root('docProps/app.xml')
    headings = next((n for n in root if local(n.tag) == 'HeadingPairs'), None)
    titles = next((n for n in root if local(n.tag) == 'TitlesOfParts'), None)
    if headings is None or titles is None:
        return
    updated_headings, updated_titles = copy.deepcopy(headings), copy.deepcopy(titles)
    variants = list(updated_headings[0])
    offset = 0
    for index in range(0, len(variants), 2):
        count = variants[index + 1][0]
        number = int(count.text)
        if variants[index][0].text == 'Slide Titles':
            require(number == len(slide_parts(original)), 'slide title metadata does not match source slide count')
            vector = updated_titles[0]
            items = list(vector)
            group = items[offset:offset + number]
            if 'move_slide' in spec:
                change = spec['move_slide']
                group.insert(change['to'] - 1, group.pop(change['from'] - 1))
            elif 'delete_slide' in spec:
                del group[spec['delete_slide'] - 1]
            else:
                group[-2:] = reversed(group[-2:])
            vector[:] = items[:offset] + group + items[offset + number:]
            vector.set('size', str(len(vector)))
            count.text = str(len(group))
            break
        offset += number
    else:
        return
    actual_root = actual.root('docProps/app.xml')
    found = [next((n for n in actual_root if n.tag == source.tag), None) for source in (headings, titles)]
    state = tuple(canonical(n) for n in found)
    require(state in (tuple(map(canonical, (headings, titles))), tuple(map(canonical, (updated_headings, updated_titles)))),
            'unexpected slide title metadata')
    for source, node in zip((headings, titles), found):
        index = list(actual_root).index(node)
        actual_root.remove(node)
        actual_root.insert(index, copy.deepcopy(source))


def check_real(task, source, revised):
    original, actual = RealPackage(source), RealPackage(revised)
    spec = copy.deepcopy(task['expected'])
    if task['format'] == 'xlsx':
        for changes in spec.get('cells', {}).values():
            for patch in changes.values():
                if 'fill' in patch.get('style', {}):
                    patch['style']['fill_other'] = ()
                    patch['style']['fill_rest'] = canonical(ET.fromstring(f'<fill xmlns="{NS["s"]}"><patternFill/></fill>'))
        real_metadata(original)
        real_metadata(actual)
        calculation = original.root('xl/workbook.xml').find('s:calcPr', NS)
        saved_calculation = actual.root('xl/workbook.xml').find('s:calcPr', NS)
        if calculation is not None and saved_calculation is not None and saved_calculation.get('fullCalcOnLoad') in ('1', 'true'):
            calculation.set('fullCalcOnLoad', saved_calculation.get('fullCalcOnLoad'))
        renames = spec.get('rename', {})
        if renames and 'docProps/app.xml' in original.parts:
            app = original.root('docProps/app.xml')
            for node in app.iter():
                if local(node.tag) == 'TitlesOfParts':
                    for item in node.iter():
                        if local(item.tag) == 'lpstr' and item.text in renames:
                            item.text = renames[item.text]
        before = workbook(original, real=True)
        cell_styles, _ = styles(original, real=True)
        links = relations(original, 'xl/workbook.xml')
        sheet_parts = {n.get('name'): links[n.get(f'{{{NS["r"]}}}id')]['target'] for n in xml(original, 'xl/workbook.xml').findall('s:sheets/s:sheet', NS)}
        for sheet in before['sheets']:
            root = xml(original, sheet_parts[sheet['name']])
            for address in spec.get('cells', {}).get(sheet['name'], {}):
                if address in sheet['cells']:
                    continue
                number = re.search('[0-9]+$', address).group()
                row = root.find(f's:sheetData/s:row[@r="{number}"]', NS)
                column = column_number(re.match('[A-Z]+', address).group())
                inherited = row.get('s', '0') if row is not None and row.get('customFormat') in ('1', 'true') else next((n.get('style', '0') for n in root.findall('s:cols/s:col', NS) if int(n.get('min')) <= column <= int(n.get('max'))), '0')
                sheet['cells'][address] = {'value': None, 'formula': None, 'style': copy.deepcopy(indexed(cell_styles, inherited, 'cell format')),
                                           'attributes': {}, 'formula_properties': {}, 'other': (), 'rich_text': None}
        wanted = expected_workbook(before, spec)
        if 'sort_range' in spec:
            change = spec['sort_range']
            sheet = next(s for s in wanted['sheets'] if s['name'] == change['sheet'])
            first, low, last, high = re.fullmatch(r'([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)', change['range']).groups()
            low, high = int(low), int(high)
            columns = range(column_number(first), column_number(last) + 1)
            source_cells = copy.deepcopy(sheet['cells'])
            order = sorted(range(low, high + 1), key=lambda row: source_cells[f'{change["key"]}{row}']['value'].casefold(), reverse=not change['ascending'])
            for destination, origin in zip(range(low, high + 1), order):
                for column in columns:
                    address = f'{column_letters(column)}{destination}'
                    previous = f'{column_letters(column)}{origin}'
                    sheet['cells'].pop(address, None)
                    if previous in source_cells:
                        sheet['cells'][address] = copy.deepcopy(source_cells[previous])
        renames = spec.get('rename', {})
        for sheet in wanted['sheets']:
            for record in sheet['cells'].values():
                record['formula'] = rename_formula(record['formula'], renames)
            for address, change in spec.get('cell_replace', {}).get(sheet['name'], {}).items():
                record = sheet['cells'][address]
                require(change['old'] in record['value'], 'cell replacement does not match fixture')
                record['value'] = record['value'].replace(change['old'], change['new'])
            for address, patch in spec.get('cells', {}).get(sheet['name'], {}).items():
                record = sheet['cells'][address]
                record.setdefault('attributes', {})
                record.setdefault('formula_properties', {})
                record.setdefault('other', ())
                record.setdefault('rich_text', None)
                if 'value' in patch:
                    require(record['rich_text'] is None, 'edited rich text cell requires explicit formatting expectation')
        found = workbook(actual, real=True)
        saved_sheets = {sheet['name']: sheet for sheet in found['sheets']}
        for sheet in wanted['sheets']:
            saved_cells = saved_sheets.get(sheet['name'], {}).get('cells', {})
            for address, record in sheet['cells'].items():
                saved = saved_cells.get(address, {})
                if (record['formula'] is not None and saved.get('formula') is not None
                        and isinstance(record['value'], Decimal) and isinstance(saved.get('value'), Decimal)
                        and math.isclose(record['value'], saved['value'], rel_tol=1e-12, abs_tol=0.0)):
                    record['value'] = saved['value']
        compare(wanted, found)
        return
    ids = real_ids(original) if task['format'] == 'docx' else {}
    all_ids = {value for _, value in ids}
    if task['format'] == 'pptx':
        real_slide_metadata(original, actual, spec)
    real_expected_package(original, spec, task['format'])
    if task['format'] == 'docx':
        compare(tracked_changes(original), tracked_changes(actual), 'tracked revisions')
        after = real_ids(actual)
        for key, count in after.items():
            require(count <= max(1, ids.get(key, 0)), f'new duplicate paragraph identity: {key}')
    compare(real_package_state(original, True, all_ids), real_package_state(actual))


def check(task, output):
    output = Path(output)
    require(output.is_file(), f'missing output file: {output.name}')
    with ZipFile(fixture(task)) as source, ZipFile(output) as revised:
        validate_package(revised)
        if task.get('suite', 'v2') == 'v3':
            check_real(task, source, revised)
            return 'requested edit and all preserved package content verified'
        if task['format'] == 'xlsx':
            compare(expected_workbook(workbook(source), task['expected']), workbook(revised))
        elif task['format'] == 'docx':
            check_document(expected_document(document(source), task['expected']), document(revised))
        else:
            check_presentation(expected_presentation(presentation(source), task['expected']), presentation(revised))
    return 'exact requested content, formulas, formatting, order, notes, and preserved content verified'


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--suite', choices=('v2', 'v3'), default='v2')
    parser.add_argument('task')
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    task = next((task for task in tasks(args.suite) if task['id'] == args.task), None)
    if task is None:
        parser.error(f'unknown {args.suite} task: {args.task}')
    try:
        print(f'PASS: {task["id"]}: {check(task, args.output)}')
    except Exception as error:
        print(f'FAIL: {task["id"]}: {error}')
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
