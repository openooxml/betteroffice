import argparse
import copy
import hashlib
import itertools
import json
import math
import os
import posixpath
import re
import subprocess
import uuid
import xml.etree.ElementTree as ET
from decimal import Decimal
from io import BytesIO
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
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
COMMENT_CONTENT_TYPES = {
    f'application/vnd.ms-word.{kind}+xml': f'application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml'
    for kind in ('comments', 'commentsExtended', 'commentsIds', 'commentsExtensible')
}


def tasks(suite='v2', tier='basic'):
    require(suite in ('v2', 'v3'), f'unknown suite: {suite}')
    require(tier in ('basic', 'complex', 'all'), f'unknown tier: {tier}')
    require(suite == 'v3' or tier == 'basic', 'complex tier requires v3')
    override = os.environ.get('FORMAT_EVAL_TASKS')
    path = Path(override) if override else ROOT / 'v3' / 'tasks.json' if suite == 'v3' else ROOT / 'tasks.json'
    basic = json.loads(path.read_text()) if tier != 'complex' else []
    complex_tasks = json.loads((ROOT / 'v3' / 'complex.json').read_text()) if tier != 'basic' else []
    return [dict(task, suite=suite) for task in basic + complex_tasks]


def verified(data, entry, name):
    if len(data) != entry['bytes'] or hashlib.sha256(data).hexdigest() != entry['sha256']:
        raise ValueError(f'fixture {name} does not match fixtures.json')
    return data


def fixture(task):
    default = ROOT / ('v3/fixtures.json' if task.get('suite', 'v2') == 'v3' else 'fixtures.json')
    manifest = json.loads(Path(os.environ.get('FORMAT_EVAL_MANIFEST') or default).read_text())
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


def prepare_input(task, data):
    for step in task.get('prepare', []):
        require(step == 'remove_sheet_protection', f'unknown prepare step: {step}')
        prepared = BytesIO()
        with ZipFile(BytesIO(data)) as source, ZipFile(prepared, 'w') as target:
            target.comment = source.comment
            for entry in source.infolist():
                content = source.read(entry)
                if re.fullmatch(r'xl/worksheets/[^/]+\.xml', entry.filename):
                    content = re.sub(rb'<sheetProtection(?=[\s/])[^>]*?/\s*>', b'', content)
                target.writestr(copy.copy(entry), content)
        data = prepared.getvalue()
    return data


def normalized_content_type(value):
    return COMMENT_CONTENT_TYPES.get(value, value)


def xml(archive, path):
    return ET.fromstring(archive.read(path))


def canonical(node):
    if node is None:
        return None
    attrs = dict(node.attrib)
    if node.tag in (f'{{{CT_NS}}}Default', f'{{{CT_NS}}}Override') and 'ContentType' in attrs:
        attrs['ContentType'] = normalized_content_type(attrs['ContentType'])
    text = node.text if node.text and node.text.strip() else None
    children = [canonical(child) for child in node]
    if node.tag in ORDER_FREE_CHILDREN:
        children.sort(key=lambda child: child[0])
    return (node.tag, tuple(sorted(attrs.items())), text, tuple(children))


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
        return normalized_content_type(overrides[0])
    extension = part.rsplit('.', 1)[-1]
    defaults = [item.attrib['ContentType'] for item in root if item.get('Extension') == extension]
    require(len(defaults) == 1, f'missing or duplicate default content type for {part}')
    return normalized_content_type(defaults[0])


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


def validate_package(archive, ignore_trash=False):
    names = archive.namelist()
    require(len(names) == len(set(names)), 'duplicate ZIP members')
    require(archive.testzip() is None, 'invalid ZIP checksum')
    for name in names:
        if ignore_trash and name.startswith('[trash]/'):
            continue
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
            if ignore_trash and item.attrib['PartName'].lstrip('/').startswith('[trash]/'):
                continue
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
        self.root_namespaces = {}
        self.text_shapes = {}

    def namelist(self):
        return list(self.parts)

    def read(self, name):
        return ET.tostring(self.roots[name]) if name in self.roots else self.parts[name]

    def root(self, name):
        if name not in self.roots:
            import io
            self.namespaces[name] = dict(item for _, item in ET.iterparse(io.BytesIO(self.parts[name]), events=('start-ns',)))
            self.root_namespaces[name] = {}
            for event, item in ET.iterparse(io.BytesIO(self.parts[name]), events=('start-ns', 'start')):
                if event == 'start':
                    break
                self.root_namespaces[name][item[0]] = item[1]
            self.roots[name] = ET.fromstring(self.parts[name])
        return self.roots[name]


def paragraph_nodes(root, prefix):
    return list(root.iter(f'{{{NS[prefix]}}}p'))


def node_text(node, prefix):
    return ''.join(n.text or '' for n in node.iter(f'{{{NS[prefix]}}}t'))


def text_middle(old, new):
    left = 0
    while left < min(len(old), len(new)) and old[left] == new[left]:
        left += 1
    right = 0
    while right < min(len(old), len(new)) - left and old[-right - 1] == new[-right - 1]:
        right += 1
    return left, old[left:len(old) - right], new[left:len(new) - right]


def replace_xml_text(paragraph, old, new, prefix, minimal=False, insertion_sides=()):
    nodes = list(paragraph.iter(f'{{{NS[prefix]}}}t'))
    text = ''.join(n.text or '' for n in nodes)
    starts = [m.start() for m in re.finditer(re.escape(old), text)]
    leading, removed, inserted = text_middle(old, new) if minimal else (0, old, new)
    for index, start in enumerate(reversed(starts)):
        start += leading
        offset = 0
        end = start + len(removed)
        if minimal and not removed:
            if not inserted:
                continue
            left_node = right_node = None
            for node in nodes:
                length = len(node.text or '')
                if offset < start <= offset + length:
                    left_node = (node, start - offset)
                if offset <= start < offset + length:
                    right_node = (node, start - offset)
                offset += length
            target = (left_node or right_node) if insertion_sides and insertion_sides[index] == 'left' else (right_node or left_node)
            require(target is not None, 'text insertion has no neighbouring run')
            node, position = target
            value = node.text or ''
            node.text = value[:position] + inserted + value[position:]
            if prefix == 'w' and (node.text[0].isspace() or node.text[-1].isspace()):
                node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            continue
        for node in nodes:
            value = node.text or ''
            left, right = max(0, start - offset), min(len(value), end - offset)
            if left < right:
                node.text = value[:left] + (inserted if offset <= start < offset + len(value) else '') + value[right:]
                if prefix == 'w' and node.text and (node.text[0] in ' \t\r\n' or node.text[-1] in ' \t\r\n'):
                    node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            offset += len(value)
    return len(starts)


def real_replace_xml_text(package, paragraph, old, new, prefix):
    if old in node_text(paragraph, prefix):
        source, shapes = package.text_shapes.get(paragraph, (copy.deepcopy(paragraph), [copy.deepcopy(paragraph)]))
        updated = {}
        for shape in shapes:
            text = node_text(shape, prefix)
            starts = [m.start() for m in re.finditer(re.escape(old), text)]
            leading, removed, inserted = text_middle(old, new)
            boundaries = set(itertools.accumulate(len(n.text or '') for n in shape.iter(f'{{{NS[prefix]}}}t'))) - {0, len(text)}
            ambiguous = [i for i, start in enumerate(reversed(starts)) if start + leading in boundaries] if not removed and inserted else []
            for choices in itertools.product(('left', 'right'), repeat=len(ambiguous)):
                sides = ['right'] * len(starts)
                for index, side in zip(ambiguous, choices):
                    sides[index] = side
                candidate = copy.deepcopy(shape)
                replace_xml_text(candidate, old, new, prefix, minimal=True, insertion_sides=sides)
                updated[tuple(n.text for n in candidate.iter(f'{{{NS[prefix]}}}t'))] = candidate
        package.text_shapes[paragraph] = (source, list(updated.values()))
    return replace_xml_text(paragraph, old, new, prefix)


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
                matches += real_replace_xml_text(package, paragraph, edit['old'], edit['new'], prefix)
        require(matches > 0, f'fixture edit does not match: {edit}')
    if 'slide_replace' in spec:
        change = spec['slide_replace']
        for name in package.namelist():
            if re.fullmatch(r'ppt/slides/slide[0-9]+.xml', name):
                for paragraph in paragraph_nodes(package.root(name), 'a'):
                    real_replace_xml_text(package, paragraph, change['old'], change['new'], 'a')
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
    def __init__(self, attributes, input_ids, paragraph=True, declared_namespaces=()):
        dict.__init__(self, attributes)
        self.input_ids = input_ids
        self.paragraph = paragraph
        self.declared_namespaces = set(declared_namespaces)

    def compare(self, actual, path):
        expected = dict(self)
        ignorable = f'{{{NS["mc"]}}}Ignorable'
        declared = getattr(actual, 'declared_namespaces', set())
        actual = dict(actual)
        if not self.paragraph:
            tokens = set(actual.get(ignorable, '').split())
            required = set(self.get(ignorable, '').split())
            require(required <= tokens and tokens - required <= declared, f'{path}: unexpected ignorable namespaces')
            if ignorable in self:
                actual[ignorable] = self[ignorable]
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


class RealTextChoices(NewParagraphAttributes):
    def __init__(self, shapes):
        self.shapes = shapes

    def compare(self, actual, path):
        for shape in self.shapes:
            try:
                compare(shape, actual, path)
                return
            except ValueError:
                pass
        raise ValueError(f'{path}: edited paragraph differs from legacy and minimal-span shapes')


class RealTextTokens(NewParagraphAttributes):
    def __init__(self, tokens, optional):
        self.tokens = tokens
        self.optional = optional

    def compare(self, actual, path):
        require(isinstance(actual, (list, tuple)), f'{path}: invalid paragraph content')
        positions = {0}
        for index, token in enumerate(self.tokens):
            following = set(positions) if index in self.optional else set()
            for position in positions:
                if position < len(actual):
                    try:
                        compare(token, actual[position], f'{path}[{position + 1}]')
                        following.add(position + 1)
                    except ValueError:
                        pass
            positions = following
            require(positions, f'{path}: paragraph content differs')
        require(len(actual) in positions, f'{path}: paragraph content count differs')


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
    source_root = package.root(name)
    root = copy.deepcopy(source_root)
    shapes = {after: package.text_shapes[before] for before, after in zip(source_root.iter(), root.iter())
              if before in package.text_shapes} if expected else {}
    namespaces = package.namespaces[name]
    links = relations(package, name)
    styles_root = package.root('word/styles.xml') if 'word/styles.xml' in package.parts and name.startswith('word/') else None
    defaults = [s.get(w('styleId')) for s in styles_root.findall('w:style', NS) if s.get(w('type')) == 'paragraph' and s.get(w('default')) in ('1', 'true', 'on')] if styles_root is not None else []
    default_style = defaults[0] if defaults else None
    def model(node, paragraph_style=None, emptied=frozenset(), edited=False):
        if node in shapes and not edited:
            source, alternatives = shapes[node]
            prefix = 'w' if node.tag == w('p') else 'a'
            run_tag, text_tag = (f'{{{NS[prefix]}}}{kind}' for kind in ('r', 't'))
            options = []
            for alternative in [node, *alternatives]:
                if node_text(alternative, prefix) != node_text(node, prefix):
                    continue
                candidate = copy.deepcopy(node)
                for target, value in zip(candidate.iter(text_tag), alternative.iter(text_tag)):
                    target.text = value.text
                    if prefix == 'w' and target.text and (target.text[0].isspace() or target.text[-1].isspace()):
                        target.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                optional = {after for before, after in zip(source.iter(run_tag), candidate.iter(run_tag))
                            if node_text(before, prefix) and not node_text(after, prefix)}
                options.append(model(candidate, paragraph_style, optional, edited=True))
            return RealTextChoices(options)
        attrs = dict(node.attrib)
        for key, value in list(attrs.items()):
            if key.startswith(f'{{{NS["r"]}}}') and value:
                require(value in links, f'missing relationship ID in {name}: {value}')
                attrs[key] = json.dumps(links[value], sort_keys=True)
        if node.tag == f'{{{NS["mc"]}}}Choice' and 'Requires' in attrs:
            attrs['Requires'] = ' '.join(sorted(namespaces.get(value, value) for value in attrs['Requires'].split()))
        if f'{{{NS["mc"]}}}Ignorable' in attrs:
            key = f'{{{NS["mc"]}}}Ignorable'
            declared = package.root_namespaces.get(name, namespaces) if node is root else namespaces
            uris = {declared.get(value, value) for value in attrs[key].split()}
            if uris:
                attrs[key] = ' '.join(sorted(uris))
            else:
                del attrs[key]
        if node is root and name.startswith('word/'):
            declared = package.root_namespaces.get(name, namespaces)
            attrs = RealIdentity(attrs, input_ids, paragraph=False,
                                 declared_namespaces=(declared[token] for token in node.get(f'{{{NS["mc"]}}}Ignorable', '').split() if token in declared))
        if node.tag == w('p'):
            props = node.find('w:pPr', NS)
            style = props.find('w:pStyle', NS) if props is not None else None
            paragraph_style = style.get(w('val')) if style is not None else default_style
            if expected:
                attrs = RealIdentity(attrs, input_ids)
        tokens = []
        optional_tokens = set()
        for child in node:
            if child.tag == w('proofErr'):
                continue
            if node.tag == w('p') and child.tag == w('pPr'):
                props = copy.deepcopy(child)
                item = props.find('w:pStyle', NS)
                if item is not None and item.get(w('val')) == default_style:
                    props.remove(item)
                if props.attrib or len(props):
                    tokens.append(model(props, paragraph_style, emptied))
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
                        content.append(('run-content', model(item, paragraph_style, emptied), props, tuple(sorted(run_attrs.items()))))
                if not content and child in emptied:
                    optional_tokens.add(len(tokens))
                tokens.extend(content or [('empty-run', props, tuple(sorted(run_attrs.items())))])
            else:
                tokens.append(model(child, paragraph_style, emptied))
        return {'tag': node.tag, 'attributes': attrs, 'text': node.text if node.text and (local(node.tag) in ('t', 'instrText', 'delText') or node.text.strip()) else None,
                'children': RealTextTokens(tokens, optional_tokens) if optional_tokens else tokens}
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


def real_reachable_parts(package):
    visited, pending = set(), ['']
    while pending:
        part = pending.pop()
        if part in visited:
            continue
        visited.add(part)
        pending.extend(link['target'] for link in relations(package, part).values()
                       if not link['external'] and link['target'] in package.parts)
    return visited - {''}


def real_deleted_orphans(package, actual, reachable):
    orphans = reachable - real_reachable_parts(package)
    removed = set()
    for name in orphans & package.parts.keys():
        if name not in actual.parts:
            require(relpath(name) not in actual.parts, f'orphan relationships retained without their part: {name}')
            removed.update((name, relpath(name)))
    for name in removed:
        package.parts.pop(name, None)
        package.roots.pop(name, None)
    types = package.root('[Content_Types].xml')
    for node in list(types):
        if node.get('PartName', '').lstrip('/') in removed:
            types.remove(node)


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
        accept_full_calc_on_load(original, actual)
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
        volatile_cells = {}
        for sheet in before['sheets']:
            root = xml(original, sheet_parts[sheet['name']])
            for cell in root.findall('s:sheetData/s:row/s:c', NS):
                formula = cell.find('s:f', NS)
                if formula is not None and formula.get('ca') == '1':
                    volatile_cells[(renames.get(sheet['name'], sheet['name']), cell.get('r'))] = (formula.text, dict(formula.attrib), cell.get('t', 'n'))
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
        saved_links = relations(actual, 'xl/workbook.xml')
        saved_parts = {n.get('name'): saved_links[n.get(f'{{{NS["r"]}}}id')]['target']
                       for n in xml(actual, 'xl/workbook.xml').findall('s:sheets/s:sheet', NS)}
        for sheet in wanted['sheets']:
            saved_cells = saved_sheets.get(sheet['name'], {}).get('cells', {})
            saved_nodes = {cell.get('r'): cell for cell in actual.root(saved_parts[sheet['name']]).findall('s:sheetData/s:row/s:c', NS)} if sheet['name'] in saved_parts else {}
            for address, record in sheet['cells'].items():
                saved = saved_cells.get(address, {})
                cell = saved_nodes.get(address)
                formula = cell.find('s:f', NS) if cell is not None else None
                if (formula is not None and (sheet['name'], address) in volatile_cells
                        and volatile_cells[(sheet['name'], address)] == (formula.text, dict(formula.attrib), cell.get('t', 'n'))
                        and record['formula'] == saved.get('formula') and record['formula_properties'] == saved.get('formula_properties')):
                    record['value'] = saved['value']
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
    reachable = real_reachable_parts(original) if task['format'] == 'pptx' and 'delete_slide' in spec else None
    real_expected_package(original, spec, task['format'])
    if reachable is not None:
        real_deleted_orphans(original, actual, reachable)
    if task['format'] == 'docx':
        compare(tracked_changes(original), tracked_changes(actual), 'tracked revisions')
        after = real_ids(actual)
        for key, count in after.items():
            require(count <= max(1, ids.get(key, 0)), f'new duplicate paragraph identity: {key}')
    compare(real_package_state(original, True, all_ids), real_package_state(actual))


COMPLEX_NS = {
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'p15': 'http://schemas.microsoft.com/office/powerpoint/2012/main',
}
NS.update(COMPLEX_NS)
COMPLEX_ORDER = {
    ('s', 'workbook'): 'fileVersion fileSharing workbookPr workbookProtection bookViews sheets functionGroups externalReferences definedNames calcPr oleSize customWorkbookViews pivotCaches smartTagPr smartTagTypes webPublishing fileRecoveryPr webPublishObjects extLst',
    ('s', 'worksheet'): 'sheetPr dimension sheetViews sheetFormatPr cols sheetData sheetCalcPr sheetProtection protectedRanges scenarios autoFilter sortState dataConsolidate customSheetViews mergeCells phoneticPr conditionalFormatting dataValidations hyperlinks printOptions pageMargins pageSetup headerFooter rowBreaks colBreaks customProperties cellWatches ignoredErrors smartTags drawing legacyDrawing legacyDrawingHF picture oleObjects controls webPublishItems tableParts extLst',
    ('s', 'styleSheet'): 'numFmts fonts fills borders cellStyleXfs cellXfs cellStyles dxfs tableStyles colors extLst',
    ('w', 'pPr'): 'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange',
    ('w', 'tcPr'): 'cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark headers cellIns cellDel cellMerge tcPrChange',
    ('w', 'sectPr'): 'headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings sectPrChange',
    ('c', 'chart'): 'title autoTitleDeleted pivotFmts view3D floor sideWall backWall plotArea legend plotVisOnly dispBlanksAs showDLblsOverMax extLst',
    ('c', 'title'): 'tx layout overlay spPr txPr extLst',
    ('c', 'ser'): 'idx order tx spPr invertIfNegative pictureOptions marker dPt dLbls trendline errBars cat val xVal yVal bubbleSize bubble3D smooth shape extLst',
}


def q(prefix, name):
    return f'{{{NS[prefix]}}}{name}'


def complex_xml_order(package):
    ranks = {q(prefix, name): {q(prefix, tag): index for index, tag in enumerate(order.split())}
             for (prefix, name), order in COMPLEX_ORDER.items()}
    for root in package.roots.values():
        for node in root.iter():
            if node.tag in ranks:
                node[:] = sorted(node, key=lambda n: ranks[node.tag].get(n.tag, 1000))


def child(parent, prefix, name, **attrs):
    node = parent.find(f'{prefix}:{name}', NS)
    if node is None:
        node = ET.SubElement(parent, q(prefix, name))
    node.attrib.update({key: str(value) for key, value in attrs.items()})
    return node


def unique(nodes, message):
    nodes = list(nodes)
    require(len(nodes) == 1, message)
    return nodes[0]


def complex_paragraph(package, change):
    return unique((p for p in package.root(change['part']).iter(w('p'))
                   if node_text(p, 'w') == change['paragraph']), 'paragraph anchor must be unique')


def fresh_id(nodes, attribute):
    return str(max([0] + [int(n.get(attribute)) for n in nodes if (n.get(attribute) or '').isdigit()]) + 1)


def add_part(package, name, root, kind):
    require(name not in package.parts, f'part already exists: {name}')
    package.parts[name] = b''
    package.roots[name] = root
    package.namespaces[name] = dict(NS)
    ET.SubElement(package.root('[Content_Types].xml'), f'{{{CT_NS}}}Override',
                  PartName='/' + name, ContentType=kind)
    return name


def new_part_name(package, folder, stem):
    number = 1
    while f'{folder}/{stem}{number}.xml' in package.parts:
        number += 1
    return f'{folder}/{stem}{number}.xml'


def add_link(package, owner, target, kind):
    path = relpath(owner)
    if path not in package.parts:
        package.parts[path] = b''
        package.roots[path] = ET.Element(q('rel', 'Relationships'))
        package.namespaces[path] = dict(NS)
    links = package.root(path)
    identity = 'rIdComplex1'
    number = 1
    while any(n.get('Id') == identity for n in links):
        number += 1
        identity = f'rIdComplex{number}'
    relation_type = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended' if kind == 'commentsExtended' else NS['r'] + '/' + kind
    ET.SubElement(links, q('rel', 'Relationship'), Id=identity, Type=relation_type,
                  Target=posixpath.relpath(target, posixpath.dirname(owner) or '.'))
    return identity


def word_run(text, template=None, deleted=False):
    run = ET.Element(w('r'))
    props = template.find('w:rPr', NS) if template is not None else None
    if props is not None:
        run.append(copy.deepcopy(props))
    node = ET.SubElement(run, w('delText' if deleted else 't'))
    node.text = text
    if text and (text[0] in ' \t\r\n' or text[-1] in ' \t\r\n'):
        node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return run


def word_span(paragraph, phrase):
    text = node_text(paragraph, 'w')
    require(phrase and text.count(phrase) == 1, 'phrase must occur exactly once')
    start, offset, selected = text.index(phrase), 0, []
    end = start + len(phrase)
    for run in list(paragraph):
        if run.tag != w('r'):
            value = node_text(run, 'w')
            require(not value or offset + len(value) <= start or offset >= end, 'phrase overlaps nested content')
            offset += len(value)
            continue
        value = node_text(run, 'w')
        low, high = max(0, start - offset), min(len(value), end - offset)
        offset += len(value)
        if low >= high:
            continue
        require(all(n.tag in (w('rPr'), w('t')) for n in run), 'phrase overlaps non-text run content')
        index = list(paragraph).index(run)
        pieces = [word_run(v, run) for v in (value[:low], value[low:high], value[high:]) if v]
        paragraph[index:index + 1] = pieces
        selected.append(pieces[1 if low else 0])
    require(selected, 'phrase has no text runs')
    return selected


def revision_text(node):
    return ''.join(n.text or '' for n in node.iter() if n.tag in (w('t'), w('delText')))


def complex_tracked(package, changes):
    for change in changes:
        root = package.root(change['part'])
        action = change['action']
        if action in ('accept', 'reject'):
            revision = unique((n for n in root.iter() if n.tag in (w('ins'), w('del'))
                               and revision_text(n) == change['text']), 'revision text must be unique')
            parent = next(p for p in root.iter() if revision in list(p))
            index = list(parent).index(revision)
            parent.remove(revision)
            keep = (action == 'accept') == (revision.tag == w('ins'))
            if keep:
                for n in revision.iter(w('delText')):
                    n.tag = w('t')
                parent[index:index] = list(revision)
            continue
        require(action in ('insert', 'delete'), 'invalid tracked action')
        paragraph = complex_paragraph(package, change)
        runs = word_span(paragraph, change['after'] if action == 'insert' else change['text'])
        identity = fresh_id((n for n in root.iter() if n.tag in TRACKED), w('id'))
        wrapper = ET.Element(w('ins' if action == 'insert' else 'del'),
                             {w('id'): identity, w('author'): change['author']})
        if action == 'insert':
            wrapper.append(word_run(change['text'], runs[-1]))
            paragraph.insert(list(paragraph).index(runs[-1]) + 1, wrapper)
        else:
            index = list(paragraph).index(runs[0])
            for run in runs:
                paragraph.remove(run)
                for n in run.iter(w('t')):
                    n.tag = w('delText')
                wrapper.append(run)
            paragraph.insert(index, wrapper)


def effective_numbering(package, paragraph):
    styles_root = package.root('word/styles.xml')
    definitions = {n.get(w('styleId')): n for n in styles_root.findall('w:style', NS)}
    result, seen = {}, set()
    def apply(props):
        num = props.find('w:numPr', NS) if props is not None else None
        if num is not None:
            result.update({local(n.tag): n.get(w('val')) for n in num})
    def inherit(identity):
        require(identity not in seen, 'cyclic numbering style')
        seen.add(identity)
        style = definitions.get(identity)
        if style is not None:
            base = style.find('w:basedOn', NS)
            if base is not None:
                inherit(base.get(w('val')))
            apply(style.find('w:pPr', NS))
    props = paragraph.find('w:pPr', NS)
    style = props.find('w:pStyle', NS) if props is not None else None
    default = next((n.get(w('styleId')) for n in definitions.values() if n.get(w('type')) == 'paragraph'
                    and n.get(w('default')) in ('1', 'true')), None)
    identity = style.get(w('val')) if style is not None else default
    inherit(identity)
    apply(props)
    if 'numId' in result and result['numId'] != '0':
        if 'ilvl' not in result:
            numbering_part = related(package, 'word/document.xml', 'numbering')
            if numbering_part:
                root = package.root(numbering_part)
                num = next((n for n in root.findall('w:num', NS) if n.get(w('numId')) == result['numId']), None)
                abstract_id = num.find('w:abstractNumId', NS).get(w('val')) if num is not None else None
                abstract = next((n for n in root.findall('w:abstractNum', NS) if n.get(w('abstractNumId')) == abstract_id), None)
                levels = [n for n in abstract.findall('w:lvl', NS) if n.find('w:pStyle', NS) is not None
                          and n.find('w:pStyle', NS).get(w('val')) == identity] if abstract is not None else []
                if levels:
                    result['ilvl'] = unique(levels, 'ambiguous numbering style level').get(w('ilvl'))
            result.setdefault('ilvl', '0')
    return result


def complex_lists(package, changes):
    for change in changes:
        root = package.root(change['part'])
        anchor = unique((p for p in root.iter(w('p')) if node_text(p, 'w') == change['after']),
                        'list anchor must be unique')
        numbering = effective_numbering(package, anchor)
        require(numbering.get('numId') not in (None, '0'), 'anchor is not a list item')
        paragraph = ET.Element(w('p'))
        props = anchor.find('w:pPr', NS)
        if props is not None:
            paragraph.append(copy.deepcopy(props))
        num = child(child(paragraph, 'w', 'pPr'), 'w', 'numPr')
        for name, value in numbering.items():
            child(num, 'w', name, **{w('val'): value})
        paragraph.append(word_run(change['text'], anchor.find('w:r', NS)))
        parent = next(p for p in root.iter() if anchor in list(p))
        parent.insert(list(parent).index(anchor) + 1, paragraph)


def complex_word_table(package, change):
    tables = package.root(change['part']).findall('.//w:tbl', NS)
    if isinstance(change['table'], str):
        return unique((table for table in tables if table.find('w:tblPr/w:tblCaption', NS) is not None
                       and table.find('w:tblPr/w:tblCaption', NS).get(w('val')) == change['table']),
                      'table caption must be unique')
    return tables[change['table'] - 1]


def complex_merge(package, changes):
    for change in changes:
        table = complex_word_table(package, change)
        r1, c1, r2, c2 = change['range']
        require(1 <= r1 <= r2 and 1 <= c1 <= c2 and (r1, c1) != (r2, c2), 'invalid merge range')
        for index, row in enumerate(table.findall('w:tr', NS), 1):
            if not r1 <= index <= r2:
                continue
            cells = row.findall('w:tc', NS)
            require(len(cells) >= c2, 'merge outside table')
            chosen = cells[c1 - 1:c2]
            require(all(n.find('w:tcPr/w:gridSpan', NS) is None and n.find('w:tcPr/w:vMerge', NS) is None
                        for n in chosen), 'merge requires unmerged source cells')
            first = chosen[0]
            props = first.find('w:tcPr', NS)
            if props is None:
                props = ET.Element(w('tcPr'))
                first.insert(0, props)
            if c2 > c1:
                child(props, 'w', 'gridSpan', **{w('val'): c2 - c1 + 1})
                widths = [n.find('w:tcPr/w:tcW', NS) for n in chosen]
                if all(n is not None and n.get(w('type')) == 'dxa' for n in widths):
                    child(props, 'w', 'tcW', **{w('type'): 'dxa', w('w'): sum(int(n.get(w('w'))) for n in widths)})
            if r2 > r1:
                child(props, 'w', 'vMerge', **{w('val'): 'restart' if index == r1 else 'continue'})
            for cell in chosen[1:]:
                first.extend(copy.deepcopy(n) for n in cell if n.tag != w('tcPr'))
                row.remove(cell)
            if index > r1:
                top = table.findall('w:tr', NS)[r1 - 1].findall('w:tc', NS)[c1 - 1]
                top.extend(copy.deepcopy(n) for n in first if n.tag != w('tcPr'))
                first[:] = [props, ET.Element(w('p'))]


def complex_sections(package, changes):
    for change in changes:
        require(change['position'] in ('before', 'after'), 'invalid section position')
        body = package.root(change['part']).find('w:body', NS)
        anchor = unique((p for p in body if p.tag == w('p') and node_text(p, 'w') == change['anchor']),
                        'section anchor must be a unique body paragraph')
        index = list(body).index(anchor) + (change['position'] == 'after')
        following = [n for block in list(body)[index:] for n in block.iter(w('sectPr'))]
        require(following, 'section has no terminating sectPr')
        current = following[0]
        previous = copy.deepcopy(current)
        if index > 0 and body[index - 1].tag == w('p'):
            paragraph = body[index - 1]
        else:
            paragraph = ET.Element(w('p'))
            body.insert(index, paragraph)
        props = paragraph.find('w:pPr', NS)
        if props is None:
            props = ET.Element(w('pPr'))
            paragraph.insert(0, props)
        require(props.find('w:sectPr', NS) is None, 'section boundary already exists')
        props.append(previous)
        child(current, 'w', 'type', **{w('val'): 'nextPage'})
        require(change['orientation'] in ('portrait', 'landscape'), 'invalid orientation')
        child(current, 'w', 'pgSz', **{w('orient'): change['orientation'], w('w'): change['width'], w('h'): change['height']})
        child(current, 'w', 'pgMar', **{w(k): v for k, v in change['margins'].items()})


def word_comments_part(package, owner):
    owner = related(package, '', 'officeDocument')
    part = related(package, owner, 'comments')
    if part is None:
        part = new_part_name(package, 'word', 'comments')
        add_part(package, part, ET.Element(w('comments')),
                 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml')
        add_link(package, owner, part, 'comments')
    return part


def comment_extensions(package, owner):
    owner = related(package, '', 'officeDocument')
    part = related(package, owner, 'commentsExtended')
    if part is None:
        part = new_part_name(package, 'word', 'commentsExtended')
        add_part(package, part, ET.Element(q('w15', 'commentsEx')),
                 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml')
        add_link(package, owner, part, 'commentsExtended')
    return package.root(part)


def comment_para_id(package, comment):
    paragraph = list(comment.iter(w('p')))[-1]
    identity = paragraph.get(q('w14', 'paraId'))
    if identity is None:
        used = {value for _, value in real_ids(package)}
        number = 1
        while f'{number:08X}' in used:
            number += 1
        identity = f'{number:08X}'
        paragraph.set(q('w14', 'paraId'), identity)
    return identity


def word_comment_body(comment):
    return '\n'.join(''.join(n.text or '' if n.tag == w('t') else '\t' if n.tag == w('tab') else '\n'
                             if n.tag in (w('br'), w('cr')) else '' for n in paragraph.iter())
                     for paragraph in comment.iter(w('p')))


def complex_word_comments(package, changes):
    for change in changes:
        owner = change['part']
        comments = package.root(word_comments_part(package, owner))
        action = change['action']
        require(action in ('add', 'reply', 'resolve'), 'invalid comment action')
        parent = None
        if action != 'add':
            parent = unique((n for n in comments if word_comment_body(n) == change['comment']),
                            'comment text must be unique')
        if action == 'resolve':
            identity = comment_para_id(package, parent)
            extensions = comment_extensions(package, owner)
            node = unique_or_new(extensions, q('w15', 'commentEx'), q('w15', 'paraId'), identity)
            node.set(q('w15', 'done'), '1')
            continue
        identity = fresh_id(comments, w('id'))
        comment = ET.SubElement(comments, w('comment'), {w('id'): identity, w('author'): change['author']})
        for line in change['text'].split('\n'):
            ET.SubElement(comment, w('p')).append(word_run(line))
        if action == 'reply':
            extensions = comment_extensions(package, owner)
            parent_id = comment_para_id(package, parent)
            unique_or_new(extensions, q('w15', 'commentEx'), q('w15', 'paraId'), parent_id)
            ET.SubElement(extensions, q('w15', 'commentEx'), {
                q('w15', 'paraId'): comment_para_id(package, comment),
                q('w15', 'paraIdParent'): parent_id})
        else:
            paragraph = complex_paragraph(package, change)
            runs = word_span(paragraph, change['phrase'])
            paragraph.insert(list(paragraph).index(runs[0]), ET.Element(w('commentRangeStart'), {w('id'): identity}))
            index = list(paragraph).index(runs[-1]) + 1
            paragraph.insert(index, ET.Element(w('commentRangeEnd'), {w('id'): identity}))
            run = ET.Element(w('r'))
            ET.SubElement(run, w('commentReference'), {w('id'): identity})
            paragraph.insert(index + 1, run)


def unique_or_new(parent, tag, attribute, value):
    matches = [n for n in parent if n.tag == tag and n.get(attribute) == value]
    require(len(matches) <= 1, 'duplicate identity')
    return matches[0] if matches else ET.SubElement(parent, tag, {attribute: value})


def sheet_map(package):
    links = relations(package, 'xl/workbook.xml')
    return {n.get('name'): links[n.get(q('r', 'id'))]['target']
            for n in package.root('xl/workbook.xml').findall('s:sheets/s:sheet', NS)}


def xlsx_cell(package, sheet, address):
    root = package.root(sheet_map(package)[sheet])
    data = root.find('s:sheetData', NS)
    number = re.search('[0-9]+$', address).group()
    row = unique_or_new(data, q('s', 'row'), 'r', number)
    node = unique_or_new(row, q('s', 'c'), 'r', address)
    return node


def set_xlsx_cell(node, value, formula=None):
    old_formula = node.find('s:f', NS)
    node[:] = [n for n in node if n.tag not in (q('s', 'f'), q('s', 'v'), q('s', 'is'))]
    node.attrib.pop('t', None)
    if formula is not None:
        attrs = dict(old_formula.attrib) if old_formula is not None else {}
        if attrs.get('t') == 'shared':
            attrs = {k: v for k, v in attrs.items() if k not in ('t', 'si', 'ref')}
        ET.SubElement(node, q('s', 'f'), attrs).text = formula
    if isinstance(value, bool):
        node.set('t', 'b')
        ET.SubElement(node, q('s', 'v')).text = '1' if value else '0'
    elif isinstance(value, str):
        node.set('t', 'str' if formula is not None else 'inlineStr')
        if formula is not None:
            ET.SubElement(node, q('s', 'v')).text = value
        else:
            ET.SubElement(ET.SubElement(node, q('s', 'is')), q('s', 't')).text = value
    elif isinstance(value, dict) and set(value) == {'error'}:
        node.set('t', 'e')
        ET.SubElement(node, q('s', 'v')).text = value['error']
    elif value is not None:
        ET.SubElement(node, q('s', 'v')).text = str(value)


def dxf_node(spec):
    node = ET.Element(q('s', 'dxf'))
    if 'font' in spec:
        font = ET.SubElement(node, q('s', 'font'))
        for name, value in spec['font'].items():
            require(name in ('bold', 'italic', 'color'), 'unsupported differential font property')
            if name == 'color':
                ET.SubElement(font, q('s', 'color'), rgb=value.upper())
            else:
                ET.SubElement(font, q('s', 'b' if name == 'bold' else 'i'), val='1' if value else '0')
    if 'fill' in spec:
        pattern = ET.SubElement(ET.SubElement(node, q('s', 'fill')), q('s', 'patternFill'), patternType='solid')
        ET.SubElement(pattern, q('s', 'fgColor'), rgb=spec['fill'].upper())
    return node


def complex_conditional(package, changes):
    styles_root = package.root(related(package, 'xl/workbook.xml', 'styles'))
    for change in changes:
        dxfs = child(styles_root, 's', 'dxfs')
        identity = len(dxfs)
        dxfs.append(dxf_node(change['dxf']))
        dxfs.set('count', str(len(dxfs)))
        root = package.root(sheet_map(package)[change['sheet']])
        priority = max([0] + [int(n.get('priority')) for n in root.findall('s:conditionalFormatting/s:cfRule', NS)]) + 1
        cf = ET.SubElement(root, q('s', 'conditionalFormatting'), sqref=change['sqref'])
        attrs = {'type': change['type'], 'priority': str(priority), 'dxfId': str(identity)}
        if 'operator' in change:
            attrs['operator'] = change['operator']
        if 'stopIfTrue' in change:
            attrs['stopIfTrue'] = '1' if change['stopIfTrue'] else '0'
        rule = ET.SubElement(cf, q('s', 'cfRule'), attrs)
        for formula in change['formulas']:
            ET.SubElement(rule, q('s', 'formula')).text = formula


def complex_charts(package, changes):
    for change in changes:
        root = package.root(change['part'])
        chart = root.find('c:chart', NS)
        title = child(chart, 'c', 'title')
        text = child(title, 'c', 'tx')
        rich = text.find('c:rich', NS)
        if rich is not None:
            drawing_text(rich, change['title'])
        else:
            text[:] = []
            rich = ET.SubElement(text, q('c', 'rich'))
            ET.SubElement(rich, q('a', 'bodyPr'))
            ET.SubElement(rich, q('a', 'lstStyle'))
            run = ET.SubElement(ET.SubElement(rich, q('a', 'p')), q('a', 'r'))
            ET.SubElement(run, q('a', 't')).text = change['title']
        series = chart.findall('.//c:ser', NS)
        require(len(series) == len(change['series']), 'chart spec must list every existing series')
        for node, spec in zip(series, change['series']):
            tx = child(node, 'c', 'tx')
            tx[:] = []
            ET.SubElement(tx, q('c', 'v')).text = spec['name']
            for key, field in (('categories', 'cat'), ('values', 'val')):
                target = node.find('c:' + field, NS)
                if target is None:
                    field = 'xVal' if key == 'categories' else 'yVal'
                    target = node.find('c:' + field, NS)
                require(target is not None, 'chart series has no data reference')
                target[:] = []
                reference = ET.SubElement(target, q('c', 'strRef' if field == 'cat' else 'numRef'))
                ET.SubElement(reference, q('c', 'f')).text = spec[key]


def rename_scoped_formula(formula, old, new, scope, context, shadows, sheets):
    pieces = re.split(r'("(?:[^"]|"")*")', formula or '')
    token = re.compile(r"(?<![\w.\\])(?:(?P<sheet>'(?:[^']|'')+'|(?:\[[^\]]+\])?[A-Za-z_\\][\w.\\]*|\[[^\]]+\])!)?(?P<name>[A-Za-z_\\][\w.\\]*)|\[[^\]]*\]")
    for index in range(0, len(pieces), 2):
        value = pieces[index]
        def replace(match):
            if match['name'] is None or match['name'].casefold() != old.casefold() or re.match(r'\s*[(\[]', value[match.end():]):
                return match.group()
            sheet = match['sheet']
            if sheet is not None and '[' in sheet:
                return match.group()
            target = sheet[1:-1].replace("''", "'") if sheet is not None and sheet.startswith("'") else sheet
            if target is not None and target.casefold() not in sheets:
                return match.group()
            target = target if target is not None else context
            target = target.casefold() if target is not None else None
            applies = target == scope.casefold() if scope is not None else target not in shadows
            return match.group()[:-len(match['name'])] + new if applies else match.group()
        pieces[index] = token.sub(replace, value)
    return ''.join(pieces)


def complex_names(package, changes):
    root = package.root('xl/workbook.xml')
    names = child(root, 's', 'definedNames')
    for change in changes:
        scope = str(change['localSheetId']) if 'localSheetId' in change else None
        require(change['action'] in ('add', 'rename'), 'invalid defined name action')
        require(not any(n.get('name', '').casefold() == change['name'].casefold() and n.get('localSheetId') == scope for n in names),
                'defined name already exists')
        if change['action'] == 'add':
            attrs = {'name': change['name']}
            if scope is not None:
                attrs['localSheetId'] = scope
            ET.SubElement(names, q('s', 'definedName'), attrs).text = change['formula']
        else:
            old = change['old']
            node = unique((n for n in names if n.get('name', '').casefold() == old.casefold()
                           and n.get('localSheetId') == scope), 'renamed name must exist uniquely')
            node.set('name', change['name'])
            sheet_names = [n.get('name') for n in root.findall('s:sheets/s:sheet', NS)]
            scope_name = sheet_names[int(scope)] if scope is not None else None
            shadows = {sheet_names[int(n.get('localSheetId'))].casefold() for n in names
                       if n.get('name', '').casefold() == old.casefold() and n.get('localSheetId') is not None}
            contexts = {part: sheet for sheet, part in sheet_map(package).items()}
            for part in package.namelist():
                if part.startswith('xl/') and part.endswith('.xml'):
                    for formula in package.root(part).iter():
                        if local(formula.tag) in ('f', 'formula', 'formula1', 'formula2', 'definedName') and formula.text:
                            context = contexts.get(part)
                            if formula.tag == q('s', 'definedName') and formula.get('localSheetId') is not None:
                                context = sheet_names[int(formula.get('localSheetId'))]
                            formula.text = rename_scoped_formula(formula.text, old, change['name'], scope_name, context, shadows,
                                                                 {name.casefold() for name in sheet_names})


def range_coordinates(value):
    first, low, last, high = re.fullmatch(r'([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)', value).groups()
    return column_number(first), int(low), column_number(last), int(high)


def complex_sort(package, changes):
    for change in changes:
        root = package.root(sheet_map(package)[change['sheet']])
        auto = child(root, 's', 'autoFilter', ref=change['ref'])
        workbook_root = package.root('xl/workbook.xml')
        index = str([n.get('name') for n in workbook_root.findall('s:sheets/s:sheet', NS)].index(change['sheet']))
        for name in workbook_root.findall('s:definedNames/s:definedName', NS):
            if name.get('name', '').casefold() == '_xlnm._filterdatabase' and name.get('localSheetId') == index and '!' in (name.text or ''):
                name.text = name.text.rsplit('!', 1)[0] + '!' + re.sub(r'([A-Z]+)(\d+)', r'$\1$\2', change['ref'])
        for parent in (root, auto):
            for state in parent.findall('s:sortState', NS):
                parent.remove(state)
        c1, r1, c2, r2 = range_coordinates(change['range'])
        require(c1 <= c2 and r1 <= r2 and change['keys'], 'invalid sort range')
        require(not set(range(r1, r2 + 1)) & set(change['fixed_rows']), 'sort range includes a fixed row')
        rows = {int(n.get('r')): n for n in root.findall('s:sheetData/s:row', NS)}
        require(all(n in rows for n in range(r1, r2 + 1)), 'sort range has missing rows')
        model = next(s for s in workbook(package, real=True)['sheets'] if s['name'] == change['sheet'])
        order = list(range(r1, r2 + 1))
        for key in reversed(change['keys']):
            require(c1 <= column_number(key['column']) <= c2, 'sort key outside range')
            def value(row):
                record = model['cells'].get(f'{key["column"]}{row}', {})
                item = record.get('value')
                require(not isinstance(item, dict), 'error cells cannot be sort keys')
                return (item is None, 1 if isinstance(item, str) else 0,
                        item.casefold() if isinstance(item, str) else item if item is not None else '')
            order.sort(key=value, reverse=not key['ascending'])
        snapshots = copy.deepcopy(rows)
        for destination, origin in zip(range(r1, r2 + 1), order):
            row = rows[destination]
            row[:] = [n for n in row if not (n.tag == q('s', 'c') and c1 <= column_number(re.match('[A-Z]+', n.get('r')).group()) <= c2)]
            for cell in snapshots[origin]:
                if cell.tag != q('s', 'c'):
                    continue
                column = re.match('[A-Z]+', cell.get('r')).group()
                if c1 <= column_number(column) <= c2:
                    old_address = cell.get('r')
                    cell.set('r', column + str(destination))
                    formula = cell.find('s:f', NS)
                    if formula is not None:
                        require(formula.get('t', 'normal') in ('normal', 'shared'), 'sort range cannot cut an array formula')
                        if formula.get('t') == 'shared':
                            for key in ('t', 'si', 'ref'):
                                formula.attrib.pop(key, None)
                        formula.text = translate_formula(model['cells'][old_address]['formula'], old_address, cell.get('r'))
                    inherited = model['cells'].get(old_address)
                    if inherited is not None and 's' not in cell.attrib:
                        source_row = snapshots[origin]
                        style = source_row.get('s', '0') if source_row.get('customFormat') in ('1', 'true') else next(
                            (n.get('style', '0') for n in root.findall('s:cols/s:col', NS)
                             if int(n.get('min')) <= column_number(column) <= int(n.get('max'))), '0')
                        cell.set('s', style)
                    row.append(cell)
            row[:] = sorted(row, key=lambda n: column_number(re.match('[A-Z]+', n.get('r')).group()) if n.tag == q('s', 'c') else 10**6)


def complex_dependents(package, changes):
    for change in changes:
        input_cell = change['input']
        node = xlsx_cell(package, input_cell['sheet'], input_cell['cell'])
        require(node.find('s:f', NS) is None, 'dependent input must be a value cell')
        set_xlsx_cell(node, input_cell['value'])
        seen = set()
        for dependent in change['cells']:
            key = (dependent['sheet'], dependent['cell'])
            require(key not in seen, 'duplicate dependent expectation')
            seen.add(key)
            node = xlsx_cell(package, *key)
            require(node.find('s:f', NS) is not None, 'dependent must have a source formula')
            set_xlsx_cell(node, dependent['value'], dependent['formula'])


def named_shape(root, name):
    return unique((node for node in root.iter() if any(identity.get('name') == name for container in node
                   if local(container.tag).startswith('nv') for identity in container if local(identity.tag) == 'cNvPr')),
                  f'shape name must be unique: {name}')


def drawing_text(body, text):
    paragraphs = body.findall('a:p', NS)
    require(paragraphs, 'text body has no paragraph template')
    template = copy.deepcopy(paragraphs[0])
    props = template.find('a:r/a:rPr', NS)
    template[:] = [n for n in template if n.tag in (q('a', 'pPr'), q('a', 'endParaRPr'))]
    for paragraph in paragraphs:
        body.remove(paragraph)
    for line in text.split('\n'):
        paragraph = copy.deepcopy(template)
        run = ET.Element(q('a', 'r'))
        if props is not None:
            run.append(copy.deepcopy(props))
        ET.SubElement(run, q('a', 't')).text = line
        paragraph.insert(1 if paragraph.find('a:pPr', NS) is not None else 0, run)
        body.append(paragraph)


def complex_layout(package, changes):
    for change in changes:
        part = slide_parts(package)[change['slide'] - 1]
        target = unique((n for n in package.namelist() if n.startswith('ppt/slideLayouts/') and n.endswith('.xml')
                         and package.root(n).find('p:cSld', NS).get('name') == change['name']), 'layout name must be unique')
        links = package.root(relpath(part))
        link = unique((n for n in links if n.get('Type').endswith('/slideLayout')), 'slide needs one layout')
        link.set('Target', posixpath.relpath(target, posixpath.dirname(part)))
        for name, mapping in change['placeholders'].items():
            shape = named_shape(package.root(part), name)
            placeholder = shape.find('p:nvSpPr/p:nvPr/p:ph', NS)
            require(placeholder is not None, 'mapped shape must be a placeholder')
            destination = named_shape(package.root(target), mapping['layout_placeholder'])
            target_ph = destination.find('p:nvSpPr/p:nvPr/p:ph', NS)
            require(target_ph is not None, 'layout mapping must identify a placeholder')
            placeholder.attrib = dict(target_ph.attrib)
            drawing_text(shape.find('p:txBody', NS), mapping['text'])


def complex_groups(package, changes):
    for change in changes:
        root = package.root(change['part'])
        group = named_shape(root, change['group'])
        require(group.tag == q('p', 'grpSp'), 'group anchor must identify a group')
        shape = named_shape(group, change['shape'])
        count = sum(replace_xml_text(p, change['old'], change['new'], 'a') for p in shape.findall('p:txBody/a:p', NS))
        require(count == 1, 'group text anchor must occur exactly once')


def fresh_table_ids(node, table):
    used = {value for item in table.iter() if local(item.tag) in ('rowId', 'colId', 'creationId') for value in item.attrib.values()}
    for item in node.iter():
        if local(item.tag) in ('rowId', 'colId', 'creationId'):
            for key, value in list(item.attrib.items()):
                if local(key) not in ('val', 'id'):
                    continue
                if value.isdigit():
                    value = str(max([0] + [int(n) for n in used if n.isdigit()]) + 1)
                else:
                    value = '{' + str(uuid.uuid4()).upper() + '}'
                item.set(key, value)
                used.add(value)


def complex_tables(package, changes):
    for change in changes:
        table = package.root(change['part']).findall('.//a:tbl', NS)[change['table'] - 1]
        rows = table.findall('a:tr', NS)
        grid = table.find('a:tblGrid', NS)
        require(change['axis'] in ('row', 'column') and change['action'] in ('insert', 'delete'), 'invalid table operation')
        index = change['index'] - 1
        count = len(rows) if change['axis'] == 'row' else len(grid)
        require(0 <= index < count + (change['action'] == 'insert'), 'table index outside table')
        require(all(not any(n.get(k) not in (None, '0', '1') for k in ('gridSpan', 'rowSpan'))
                    and n.get('hMerge', '0') in ('0', 'false') and n.get('vMerge', '0') in ('0', 'false')
                    for row in rows for n in row.findall('a:tc', NS)), 'table structure requires unmerged cells')
        if change['axis'] == 'row':
            if change['action'] == 'delete':
                require(len(rows) > 1, 'cannot delete the only row')
                table.remove(rows[index])
            else:
                new = copy.deepcopy(rows[change['template'] - 1])
                fresh_table_ids(new, table)
                cells = new.findall('a:tc', NS)
                require(len(cells) == len(change['texts']), 'inserted row text count differs')
                for cell, text in zip(cells, change['texts']):
                    drawing_text(cell.find('a:txBody', NS), text)
                table.insert(list(table).index(rows[index]) if index < len(rows) else len(table), new)
        elif change['action'] == 'delete':
            require(len(grid) > 1, 'cannot delete the only column')
            grid.remove(grid[index])
            for row in rows:
                row.remove(row.findall('a:tc', NS)[index])
        else:
            template = change['template'] - 1
            new_column = copy.deepcopy(grid[template])
            fresh_table_ids(new_column, table)
            grid.insert(index, new_column)
            require(len(rows) == len(change['texts']), 'inserted column text count differs')
            for row, text in zip(rows, change['texts']):
                cells = row.findall('a:tc', NS)
                new = copy.deepcopy(cells[template])
                fresh_table_ids(new, table)
                drawing_text(new.find('a:txBody', NS), text)
                row.insert(list(row).index(cells[index]) if index < len(cells) else len(row), new)


def clone_part(package, source, folder, stem):
    name = new_part_name(package, folder, stem)
    add_part(package, name, copy.deepcopy(package.root(source)), content_type(package, source))
    package.namespaces[name] = dict(package.namespaces.get(source, {}))
    links = relpath(source)
    if links in package.parts:
        path = relpath(name)
        package.parts[path] = b''
        package.roots[path] = copy.deepcopy(package.root(links))
        package.namespaces[path] = dict(package.namespaces.get(links, {}))
        source_links = relations(package, source)
        for link in package.roots[path]:
            record = source_links[link.get('Id')]
            if not record['external']:
                link.set('Target', posixpath.relpath(record['target'], folder))
    return name


def complex_duplicates(package, changes):
    for change in changes:
        slides = slide_parts(package)
        source = slides[change['slide'] - 1]
        name = clone_part(package, source, 'ppt/slides', 'slide')
        notes = related(package, source, 'notesSlide')
        if notes is not None:
            target = clone_part(package, notes, 'ppt/notesSlides', 'notesSlide')
            for link in package.root(relpath(name)):
                if link.get('Type').endswith('/notesSlide'):
                    link.set('Target', posixpath.relpath(target, 'ppt/slides'))
            for link in package.root(relpath(target)):
                if link.get('Type').endswith('/slide'):
                    link.set('Target', posixpath.relpath(name, 'ppt/notesSlides'))
        comments = related(package, source, 'comments')
        if comments is not None:
            target = clone_part(package, comments, 'ppt/comments', 'comment')
            for link in package.root(relpath(name)):
                if link.get('Type').endswith('/comments'):
                    link.set('Target', posixpath.relpath(target, 'ppt/slides'))
        identity = add_link(package, 'ppt/presentation.xml', name, 'slide')
        listing = package.root('ppt/presentation.xml').find('p:sldIdLst', NS)
        entry = ET.Element(q('p', 'sldId'), {'id': fresh_id(listing, 'id'), q('r', 'id'): identity})
        entry.set('id', str(max(256, int(entry.get('id')))))
        require(0 <= change['after'] <= len(slides), 'duplicate insertion outside presentation')
        listing.insert(change['after'], entry)


def complex_slide_comments(package, changes):
    for change in changes:
        part = slide_parts(package)[change['slide'] - 1]
        comments_part = related(package, part, 'comments')
        require(comments_part is not None, 'slide has no existing comments')
        comments = package.root(comments_part)
        parent = unique((n for n in comments if n.findtext('p:text', namespaces=NS) == change['comment']),
                        'slide comment text must be unique')
        authors_part = related(package, 'ppt/presentation.xml', 'commentAuthors')
        require(authors_part is not None, 'presentation has no comment authors')
        authors = package.root(authors_part)
        author = next((n for n in authors if n.get('name') == change['author']), None)
        if author is None:
            author = ET.SubElement(authors, q('p', 'cmAuthor'), id=fresh_id(authors, 'id'),
                                   name=change['author'], initials='', lastIdx='0', clrIdx='0')
        index = int(author.get('lastIdx', '0')) + 1
        author.set('lastIdx', str(index))
        comment = ET.SubElement(comments, q('p', 'cm'), authorId=author.get('id'), idx=str(index), dt='2026-01-01T00:00:00Z')
        ET.SubElement(comment, q('p', 'pos'), x='0', y='0')
        ET.SubElement(comment, q('p', 'text')).text = change['text']
        extension = ET.SubElement(ET.SubElement(comment, q('p', 'extLst')), q('p', 'ext'),
                                  uri='{C676402C-5697-4E1C-873F-D02D1690AC5C}')
        ET.SubElement(ET.SubElement(extension, q('p15', 'threadingInfo')), q('p15', 'parentCm'),
                      authorId=parent.get('authorId'), idx=parent.get('idx'))


def complex_expected_package(package, spec, fmt):
    allowed = {'docx': {'tracked_changes', 'comments', 'list_items', 'merge_cells', 'section_setup'},
               'xlsx': {'conditional_formats', 'chart', 'defined_names', 'autofilter_sort', 'dependents'},
               'pptx': {'layout', 'group_edits', 'table_structure', 'duplicate_slide', 'comments'}}
    require(spec and spec.keys() <= allowed[fmt], 'unknown or empty complex expected op')
    operations = {'docx': (('tracked_changes', complex_tracked), ('comments', complex_word_comments),
                           ('list_items', complex_lists), ('merge_cells', complex_merge), ('section_setup', complex_sections)),
                  'xlsx': (('conditional_formats', complex_conditional), ('chart', complex_charts), ('defined_names', complex_names),
                           ('autofilter_sort', complex_sort), ('dependents', complex_dependents)),
                  'pptx': (('layout', complex_layout), ('group_edits', complex_groups), ('table_structure', complex_tables),
                           ('duplicate_slide', complex_duplicates), ('comments', complex_slide_comments))}
    for name, operation in operations[fmt]:
        if name in spec:
            require(isinstance(spec[name], list) and spec[name], f'{name} must be a nonempty array')
            for change in spec[name]:
                complex_shape(fmt, name, change)
            operation(package, spec[name])


def complex_shape(fmt, name, change):
    require(isinstance(change, dict), f'{name} record must be an object')
    shapes = {
        'tracked_changes': {'insert': ('action part paragraph after text author', ''),
                            'delete': ('action part paragraph text author', ''),
                            'accept': ('action part text', ''), 'reject': ('action part text', '')},
        'list_items': ('part after text', ''), 'merge_cells': ('part table range', ''),
        'section_setup': ('part anchor position orientation width height margins', ''),
        'conditional_formats': ('sheet sqref type formulas dxf', 'operator stopIfTrue'),
        'chart': ('part title series', ''),
        'defined_names': {'add': ('action name formula', 'localSheetId'), 'rename': ('action old name', 'localSheetId')},
        'autofilter_sort': ('sheet ref range keys fixed_rows', ''), 'dependents': ('input cells', ''),
        'layout': ('slide name placeholders', ''), 'group_edits': ('part group shape old new', ''),
        'table_structure': {'insert': ('part table axis action index template texts', ''),
                            'delete': ('part table axis action index', '')},
        'duplicate_slide': ('slide after', ''),
        'comments': {'add': ('action part paragraph phrase author text', ''),
                     'reply': ('action part comment author text', ''), 'resolve': ('action part comment', '')}
                    if fmt == 'docx' else ('slide comment author text', '')}
    shape = shapes[name]
    if isinstance(shape, dict):
        require(change.get('action') in shape, f'{name}: invalid action')
        shape = shape[change['action']]
    required, optional = (set(value.split()) for value in shape)
    require(required <= change.keys() <= required | optional, f'{name}: expected fields {sorted(required)}, optional {sorted(optional)}')
    if name == 'section_setup':
        require(set(change['margins']) == {'top', 'bottom', 'left', 'right', 'header', 'footer', 'gutter'}, 'section margins must list all seven values')
    if name == 'conditional_formats':
        require(change['dxf'].keys() <= {'fill', 'font'}, 'unknown differential format field')
        require(isinstance(change['formulas'], list), 'conditional formulas must be an array')
    if name == 'dependents':
        require(set(change['input']) == {'sheet', 'cell', 'value'}, 'invalid dependent input shape')
        require(all(set(n) == {'sheet', 'cell', 'formula', 'value'} for n in change['cells']), 'invalid dependent cell shape')
    if name == 'chart':
        require(all(set(n) == {'name', 'categories', 'values'} for n in change['series']), 'invalid chart series shape')
    if name == 'layout':
        require(all(set(n) == {'layout_placeholder', 'text'} for n in change['placeholders'].values()), 'invalid placeholder mapping shape')


def complex_integrity(package, original):
    from collections import Counter
    for name in package.namelist():
        if name == '[Content_Types].xml':
            continue
        content_type(package, name)
        if name.endswith('.rels'):
            continue
        if not name.endswith('.xml'):
            continue
        root = package.root(name)
        links = relations(package, name)
        for node in root.iter():
            for key, value in node.attrib.items():
                if key.startswith('{' + NS['r'] + '}') and value:
                    require(value in links, f'dangling relationship ID in {name}: {value}')
        if name.startswith('ppt/slides/'):
            ids = [n.get('id') for n in root.iter(q('p', 'cNvPr'))]
            before = Counter(n.get('id') for n in original.root(name).iter(q('p', 'cNvPr'))) if name in original.parts else {}
            require(None not in ids and all(count <= max(1, before.get(identity, 0)) for identity, count in Counter(ids).items()),
                    f'duplicate shape ID in {name}')
        if root.tag == w('comments'):
            ids = [n.get(w('id')) for n in root]
            require(None not in ids and len(ids) == len(set(ids)), 'duplicate Word comment ID')
        if name.startswith('word/'):
            before = Counter(n.get(w('id')) for n in original.root(name).iter() if n.tag in TRACKED) if name in original.parts else {}
            after = Counter(n.get(w('id')) for n in root.iter() if n.tag in TRACKED)
            for identity, count in after.items():
                require(identity is not None and count <= max(1, before.get(identity, 0)), 'duplicate or missing revision ID')
            for revision in root.iter(w('del')):
                require(not any(n.tag == w('t') for n in revision.iter()), 'deleted revision text must use delText')
            for text in root.iter(w('t')):
                value = text.text or ''
                if value and (value[0] in ' \t\r\n' or value[-1] in ' \t\r\n'):
                    require(text.get('{http://www.w3.org/XML/1998/namespace}space') == 'preserve', 'significant Word whitespace lacks xml:space')
        if root.tag == q('p', 'cmLst'):
            ids = [(n.get('authorId'), n.get('idx')) for n in root]
            require(len(ids) == len(set(ids)), 'duplicate slide comment ID')
    if 'word/document.xml' in package.parts:
        comments = related(package, 'word/document.xml', 'comments')
        ids = {n.get(w('id')) for n in package.root(comments)} if comments else set()
        for name in package.namelist():
            if name.startswith('word/') and name.endswith('.xml'):
                for node in package.root(name).iter():
                    if node.tag in (w('commentRangeStart'), w('commentRangeEnd'), w('commentReference')):
                        require(node.get(w('id')) in ids, 'dangling Word comment ID')
        before, after = real_ids(original), real_ids(package)
        for key, count in after.items():
            require(count <= max(1, before.get(key, 0)), f'new duplicate paragraph identity: {key}')
    if 'ppt/presentation.xml' in package.parts:
        listing = package.root('ppt/presentation.xml').find('p:sldIdLst', NS)
        ids = [n.get('id') for n in listing]
        parts = slide_parts(package)
        require(len(ids) == len(set(ids)) and len(parts) == len(set(parts)), 'duplicate slide identity')
        notes_seen = set()
        for part in parts:
            notes = related(package, part, 'notesSlide')
            if notes:
                require(notes not in notes_seen, 'slides share a notes part')
                notes_seen.add(notes)
                require(related(package, notes, 'slide') == part, 'notes point to the wrong slide')
            comments = related(package, part, 'comments')
            if comments:
                authors = related(package, 'ppt/presentation.xml', 'commentAuthors')
                require(authors is not None, 'missing slide comment authors')
                author_ids = [n.get('id') for n in package.root(authors)]
                require(len(author_ids) == len(set(author_ids)), 'duplicate slide comment author')
                identities = {(n.get('authorId'), n.get('idx')) for n in package.root(comments)}
                for node in package.root(comments):
                    require(node.get('authorId') in author_ids, 'dangling slide comment author')
                    parent = node.find('.//p15:threadingInfo/p15:parentCm', NS)
                    if parent is not None:
                        require((parent.get('authorId'), parent.get('idx')) in identities, 'dangling slide comment parent')


def move_complex_parts(package, mapping):
    mapping = {old: new for old, new in mapping.items() if old != new}
    if not mapping:
        return
    for name in list(package.namelist()):
        if not name.endswith('.rels'):
            continue
        folder, filename = posixpath.split(name)
        owner = posixpath.join(posixpath.dirname(folder), filename[:-5]) if name != '_rels/.rels' else ''
        records = relations(package, owner)
        for node in package.root(name):
            record = records[node.get('Id')]
            if not record['external']:
                node.set('Target', '/' + mapping.get(record['target'], record['target']))
    for node in package.root('[Content_Types].xml'):
        part = node.get('PartName', '').lstrip('/')
        if part in mapping:
            node.set('PartName', '/' + mapping[part])
    for old, new in list(mapping.items()):
        for source, target in ((old, new), (relpath(old), relpath(new))):
            if source in package.parts:
                require(target not in package.parts or target == source, 'new part name collides with an existing part')
                package.parts[target] = package.parts.pop(source)
                if source in package.roots:
                    package.roots[target] = package.roots.pop(source)
                    package.namespaces[target] = package.namespaces.pop(source, {})


def remove_complex_part(package, part):
    for name in package.namelist():
        if name.endswith('.rels'):
            folder, filename = posixpath.split(name)
            owner = posixpath.join(posixpath.dirname(folder), filename[:-5]) if name != '_rels/.rels' else ''
            links = relations(package, owner)
            root = package.root(name)
            root[:] = [n for n in root if links[n.get('Id')]['external'] or links[n.get('Id')]['target'] != part]
    types = package.root('[Content_Types].xml')
    types[:] = [n for n in types if n.get('PartName', '').lstrip('/') != part]
    for name in (part, relpath(part)):
        package.parts.pop(name, None)
        package.roots.pop(name, None)
        package.namespaces.pop(name, None)


def normalize_comment_metadata(package, original):
    owner = related(package, '', 'officeDocument')
    comments = related(package, owner, 'comments')
    if comments is None:
        return
    old_part = related(original, owner, 'comments')
    old_paras = {p.get(q('w14', 'paraId')) for p in original.root(old_part).iter(w('p'))} if old_part else set()
    para_ids = {p.get(q('w14', 'paraId')) for p in package.root(comments).iter(w('p')) if p.get(q('w14', 'paraId'))}
    new_paras = para_ids - old_paras
    durable_ids = set()
    for kind in ('commentsIds', 'commentsExtensible', 'commentsExtended'):
        part = related(package, owner, kind)
        if part is None:
            continue
        root = package.root(part)
        old_target = related(original, owner, kind)
        originals = {canonical(n) for n in original.root(old_target)} if old_target else set()
        keep = []
        seen = set()
        for node in root:
            attrs = {local(k): v for k, v in node.attrib.items()}
            identity = attrs.get('paraId') if kind != 'commentsExtensible' else attrs.get('durableId')
            require(identity not in seen, f'duplicate {kind} identity')
            seen.add(identity)
            if kind in ('commentsIds', 'commentsExtended'):
                require(identity in para_ids, f'dangling {kind} paragraph ID')
            if canonical(node) in originals:
                keep.append(node)
            elif kind == 'commentsIds' and identity in new_paras:
                require(attrs.get('durableId'), 'missing comment durable identity')
                durable_ids.add(attrs['durableId'])
            elif kind == 'commentsExtensible' and identity in durable_ids:
                continue
            elif kind == 'commentsExtended' and set(attrs) <= {'paraId', 'done'} and attrs.get('done', '0') in ('0', 'false'):
                continue
            else:
                keep.append(node)
        root[:] = keep
        if not keep and old_target is None:
            remove_complex_part(package, part)


def align_complex_parts(expected, actual, original, fmt):
    mapping = {}
    aliases = {}
    owners = [('word/document.xml', 'word/document.xml')] if fmt == 'docx' else []
    if fmt == 'pptx':
        expected_slides, actual_slides = slide_parts(expected), slide_parts(actual)
        require(len(expected_slides) == len(actual_slides), 'slide count differs')
        original_slides = set(slide_parts(original))
        for before, after, entry, found in zip(expected_slides, actual_slides,
                                             expected.root('ppt/presentation.xml').find('p:sldIdLst', NS),
                                             actual.root('ppt/presentation.xml').find('p:sldIdLst', NS)):
            if before not in original_slides:
                require(after not in original_slides, 'duplicate must use a new slide part')
                mapping[after] = before
                entry.set('id', found.get('id'))
            owners.append((before, after))
    for before, after in owners:
        for kind in ('comments', 'commentsExtended', 'notesSlide'):
            target, found = related(expected, before, kind), related(actual, after, kind)
            require(bool(target) == bool(found), f'{kind} relationship differs')
            if target and target not in original.parts:
                require(found not in original.parts, f'{kind} must use a new part')
                mapping[found] = target
        expected_images = [n for n in relations(expected, before).values() if n['type'] == 'image' and not n['external']]
        actual_images = [n for n in relations(actual, after).values() if n['type'] == 'image' and not n['external']]
        for found in actual_images:
            if found['target'] in original.parts:
                continue
            matches = [n['target'] for n in expected_images if actual.read(found['target']) == expected.read(n['target'])
                       and content_type(actual, found['target']) == content_type(expected, n['target'])]
            if matches:
                aliases[found['target']] = matches[0]
    if aliases:
        for name in actual.namelist():
            if name.endswith('.rels'):
                folder, filename = posixpath.split(name)
                owner = posixpath.join(posixpath.dirname(folder), filename[:-5]) if name != '_rels/.rels' else ''
                links = relations(actual, owner)
                for node in actual.root(name):
                    link = links[node.get('Id')]
                    if not link['external'] and link['target'] in aliases:
                        node.set('Target', '/' + aliases[link['target']])
        for part in aliases:
            remove_complex_part(actual, part)
    move_complex_parts(actual, mapping)


def space_run(template, text, tag):
    run = ET.Element(w('r'), dict(template.attrib))
    run.extend(copy.deepcopy(n) for n in template if n.tag == w('rPr'))
    ET.SubElement(run, tag, {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = text
    return run


def boundary_text(run, space, from_end):
    items = [n for n in run if n.tag != w('rPr')] if run is not None and run.tag == w('r') else []
    edge = items[-1 if from_end else 0] if items else None
    value = edge.text or '' if edge is not None else ''
    return edge if edge is not None and (value.endswith(space) if from_end else value.startswith(space)) else None


def take_space(owner, run, edge, space, from_end):
    edge.text = edge.text[:-len(space)] if from_end else edge.text[len(space):]
    if not edge.text:
        run.remove(edge)
    if all(n.tag == w('rPr') for n in run):
        owner.remove(run)


def same_run_format(left, right):
    return dict(left.attrib) == dict(right.attrib) and [canonical(n) for n in left if n.tag == w('rPr')] == [
        canonical(n) for n in right if n.tag == w('rPr')]


def align_revision_space(package, change):
    text = change['text']
    if not text.strip(' ') or text.startswith(' ') == text.endswith(' '):
        return
    leading = text.startswith(' ')
    core = text.strip(' ')
    space = text[:len(text) - len(core)] if leading else text[len(core):]
    kind = w('ins') if change['action'] == 'insert' else w('del')
    root = package.root(change['part'])
    revisions = [n for n in root.iter(kind) if n.get(w('author')) == change['author']]
    shifted = [n for n in revisions if revision_text(n) == (core + space if leading else space + core)]
    if any(revision_text(n) == text for n in revisions) or len(shifted) != 1 or not len(shifted[0]):
        return
    revision = shifted[0]
    parent = next(p for p in root.iter() if revision in list(p))
    index = list(parent).index(revision)
    neighbour = parent[index - 1] if leading and index else parent[index + 1] if not leading and index + 1 < len(parent) else None
    far = revision[-1 if leading else 0]
    outer, inner = boundary_text(neighbour, space, leading), boundary_text(far, space, leading)
    if outer is None or outer.tag != w('t') or inner is None or inner.tag != (w('t') if kind == w('ins') else w('delText')):
        return
    if kind == w('del'):
        moved_in, moved_out = space_run(neighbour, space, w('delText')), space_run(far, space, w('t'))
    else:
        moved_in, moved_out = space_run(far, space, w('t')), space_run(neighbour, space, w('t'))
    take_space(parent, neighbour, outer, space, leading)
    take_space(revision, far, inner, space, leading)
    near = revision[0 if leading else -1] if len(revision) else None
    edge = boundary_text(near, '', not leading) if near is not None else None
    if kind == w('del') and edge is not None and edge.tag == w('delText') and same_run_format(near, moved_in):
        edge.text = space + edge.text if leading else edge.text + space
        edge.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    else:
        revision.insert(0 if leading else len(revision), moved_in)
    index = list(parent).index(revision)
    parent.insert(index + 1 if leading else index, moved_out)


def comment_markers(root, identity):
    position, markers = 0, {}
    stops = {stop for tabs in root.iter(w('tabs')) for stop in tabs}
    for item in root.iter():
        if item.tag in (w('commentRangeStart'), w('commentRangeEnd'), w('commentReference')) and item.get(w('id')) == identity:
            require(local(item.tag) not in markers, 'duplicate comment anchor marker')
            markers[local(item.tag)] = position
        if item.tag in (w('t'), w('delText')):
            position += len(item.text or '')
        elif item.tag in (w('br'), w('cr'), w('p')) or (item.tag == w('tab') and item not in stops):
            position += 1
    return markers


def remove_comment_anchors(root, scope, identity):
    reference_runs = {run for run in scope.iter(w('r')) if any(n.tag == w('commentReference')
                      and n.get(w('id')) == identity for n in run)
                      and all(n.tag in (w('rPr'), w('commentReference')) for n in run)}
    for parent in list(root.iter()):
        for node in list(parent):
            if node in reference_runs or (node.tag in (w('commentRangeStart'), w('commentRangeEnd'), w('commentReference')) and node.get(w('id')) == identity):
                parent.remove(node)
    return reference_runs


def normalize_complex_word(package, spec, original):
    for change in spec.get('tracked_changes', []):
        if change['action'] in ('insert', 'delete'):
            align_revision_space(package, change)
    support_styles = set()
    def remember_styles(node):
        support_styles.update(n.get(w('val')) for n in node.iter() if n.tag in (w('pStyle'), w('rStyle')))
    for change in spec.get('merge_cells', []):
        table = complex_word_table(package, change)
        for row in table.findall('w:tr', NS):
            first = None
            for cell in list(row.findall('w:tc', NS)):
                props = cell.find('w:tcPr', NS)
                merge = props.find('w:hMerge', NS) if props is not None else None
                if merge is None:
                    first = None
                    continue
                props.remove(merge)
                if merge.get(w('val')) == 'restart':
                    first = cell
                else:
                    require(first is not None, 'horizontal merge has no start')
                    target = first.find('w:tcPr', NS)
                    span = target.find('w:gridSpan', NS)
                    previous = int(span.get(w('val'))) if span is not None else 1
                    increment = props.find('w:gridSpan', NS)
                    child(target, 'w', 'gridSpan', **{w('val'): previous + (int(increment.get(w('val'))) if increment is not None else 1)})
                    left, right = target.find('w:tcW', NS), props.find('w:tcW', NS)
                    if left is not None and right is not None and left.get(w('type')) == right.get(w('type')) == 'dxa':
                        left.set(w('w'), str(int(left.get(w('w'))) + int(right.get(w('w')))))
                    first.extend(copy.deepcopy(n) for n in cell if n.tag != w('tcPr'))
                    row.remove(cell)
    list_paragraphs = set()
    if spec.get('section_setup'):
        body = package.root('word/document.xml').find('w:body', NS)
        inherited = {}
        for block in body:
            section = block if block.tag == w('sectPr') else block.find('w:pPr/w:sectPr', NS) if block.tag == w('p') else None
            if section is None:
                continue
            for node in list(section):
                if node.tag in (w('headerReference'), w('footerReference')):
                    inherited[(node.tag, node.get(w('type'), 'default'))] = copy.deepcopy(node)
                    section.remove(node)
            section.extend(copy.deepcopy(n) for _, n in sorted(inherited.items()))
    for change in spec.get('list_items', []):
        paragraph = unique((p for p in package.root(change['part']).iter(w('p')) if node_text(p, 'w') == change['text']),
                           'inserted list item must be unique')
        numbering = effective_numbering(package, paragraph)
        require(numbering.get('numId') not in (None, '0'), 'inserted paragraph has no numbering')
        remember_styles(paragraph)
        list_paragraphs.add(paragraph)
    comment_part = related(package, related(package, '', 'officeDocument'), 'comments')
    for change in spec.get('comments', []):
        if change['action'] != 'reply':
            continue
        comments = package.root(comment_part)
        reply = unique((n for n in comments if word_comment_body(n) == change['text']
                        and n.get(w('author')) == change['author']), 'requested comment reply missing')
        parent = unique((n for n in comments if word_comment_body(n) == change['comment']), 'comment text must be unique')
        for name in package.namelist():
            if not name.startswith('word/') or not name.endswith('.xml') or name == comment_part:
                continue
            root = package.root(name)
            markers = comment_markers(root, reply.get(w('id')))
            if markers:
                require(markers == comment_markers(root, parent.get(w('id'))), 'comment reply anchors differ from its parent')
                for run in remove_comment_anchors(root, root, reply.get(w('id'))):
                    remember_styles(run)
    for change in spec.get('comments', []):
        if change['action'] != 'add':
            continue
        part = comment_part
        comment = unique((n for n in package.root(part) if word_comment_body(n) == change['text']
                          and n.get(w('author')) == change['author']), 'requested comment missing')
        identity = comment.get(w('id'))
        paragraph = complex_paragraph(package, change)
        expected_start = change['paragraph'].index(change['phrase'])
        expected_end = expected_start + len(change['phrase'])
        position, paragraph_start, markers = 0, None, {}
        root = package.root(change['part'])
        for item in root.iter():
            if item is paragraph:
                paragraph_start = position
            if item.tag in (w('commentRangeStart'), w('commentRangeEnd'), w('commentReference')) and item.get(w('id')) == identity:
                require(local(item.tag) not in markers, 'duplicate comment anchor marker')
                markers[local(item.tag)] = position
            if item.tag == w('t'):
                position += len(item.text or '')
        expected_start += paragraph_start
        expected_end += paragraph_start
        require(markers.get('commentReference') in (expected_start, expected_end), 'comment reference is off the requested phrase')
        if 'commentRangeStart' in markers or 'commentRangeEnd' in markers:
            require(markers.get('commentRangeStart') == expected_start and markers.get('commentRangeEnd') == expected_end,
                    'comment range differs from requested phrase')
        for run in remove_comment_anchors(root, paragraph, identity):
            remember_styles(run)
    para_mapping = {}
    comment_mapping = {}
    comments_parts = [name for name in package.namelist() if name.startswith('word/') and name.endswith('.xml')
                      and package.root(name).tag == w('comments')]
    for part in comments_parts:
        root = package.root(part)
        ordered = sorted(root, key=lambda n: (word_comment_body(n), n.get(w('author'), ''), n.get(w('id'), '')))
        for index, node in enumerate(ordered):
            comment_mapping[node.get(w('id'))] = str(index)
            node.set(w('id'), str(index))
            old_part = related(original, 'word/document.xml', 'comments')
            existing = next((n for n in original.root(old_part) if word_comment_body(n) == word_comment_body(node)
                             and n.get(w('author')) == node.get(w('author'))), None) if old_part else None
            if existing is None:
                node.attrib.pop(w('date'), None)
                node.attrib.pop(w('initials'), None)
            for paragraph in node.iter(w('p')):
                identity = paragraph.get(q('w14', 'paraId'))
                if identity:
                    para_mapping[identity] = f'{index + 0x60000000:08X}'
                    paragraph.set(q('w14', 'paraId'), para_mapping[identity])
            if existing is None:
                remember_styles(node)
                text = word_comment_body(node)
                paragraph_ids = [p.get(q('w14', 'paraId')) for p in node.iter(w('p')) if p.get(q('w14', 'paraId'))]
                node[:] = []
                paragraph = ET.SubElement(node, w('p'))
                if paragraph_ids:
                    paragraph.set(q('w14', 'paraId'), paragraph_ids[-1])
                paragraph.append(word_run(text))
        root[:] = ordered
    for name in package.namelist():
        if not name.startswith('word/') or not name.endswith('.xml'):
            continue
        root = package.root(name)
        for node in root.iter():
            if node.tag in (w('commentRangeStart'), w('commentRangeEnd'), w('commentReference')):
                node.set(w('id'), comment_mapping[node.get(w('id'))])
            if node.tag == q('w15', 'commentEx'):
                for key in ('paraId', 'paraIdParent'):
                    attr = q('w15', key)
                    if attr in node.attrib:
                        require(node.get(attr) in para_mapping, 'dangling comment paragraph ID')
                        node.set(attr, para_mapping[node.get(attr)])
                if node.get(q('w15', 'done'), '0') in ('0', 'false'):
                    node.attrib.pop(q('w15', 'done'), None)
                elif node.get(q('w15', 'done')) == 'true':
                    node.set(q('w15', 'done'), '1')
            if node.tag == w('pgSz') and node.get(w('orient')) is None:
                node.set(w('orient'), 'portrait')
            if node.tag == w('vMerge') and node.get(w('val')) in (None, 'continue'):
                node.set(w('val'), 'continue')
            if node.tag == w('p'):
                numbering = effective_numbering(package, node)
                if node in list_paragraphs:
                    styles_root = package.root('word/styles.xml')
                    definitions = {n.get(w('styleId')): n for n in styles_root.findall('w:style', NS)}
                    props = node.find('w:pPr', NS)
                    style = props.find('w:pStyle', NS) if props is not None else None
                    identity = style.get(w('val')) if style is not None else next(
                        (n.get(w('styleId')) for n in definitions.values() if n.get(w('default')) in ('1', 'true')
                         and n.get(w('type')) == 'paragraph'), None)
                    for run in node.findall('w:r', NS):
                        run.set('complexRunStyle', json.dumps(effective_run_properties(run, styles_root, identity)))
                        run_props = run.find('w:rPr', NS)
                        if run_props is not None:
                            run.remove(run_props)
                    effective = {}
                    def inherit(identity, seen):
                        require(identity not in seen, 'cyclic paragraph style')
                        definition = definitions.get(identity)
                        if definition is not None:
                            base = definition.find('w:basedOn', NS)
                            if base is not None:
                                inherit(base.get(w('val')), seen | {identity})
                            ppr = definition.find('w:pPr', NS)
                            if ppr is not None:
                                effective.update({n.tag: copy.deepcopy(n) for n in ppr if n.tag not in (w('pStyle'), w('numPr'))})
                    defaults = styles_root.find('w:docDefaults/w:pPrDefault/w:pPr', NS)
                    if defaults is not None:
                        effective.update({n.tag: copy.deepcopy(n) for n in defaults})
                    inherit(identity, set())
                    if props is None:
                        props = ET.Element(w('pPr'))
                        node.insert(0, props)
                    effective.update({n.tag: copy.deepcopy(n) for n in props if n.tag not in (w('pStyle'), w('numPr'))})
                    props[:] = list(effective.values())
                if numbering.get('numId') not in (None, '0'):
                    props = node.find('w:pPr', NS)
                    if props is None:
                        props = ET.Element(w('pPr'))
                        node.insert(0, props)
                    num = child(props, 'w', 'numPr')
                    num[:] = [ET.Element(w(key), {w('val'): value}) for key, value in sorted(numbering.items())]
            if node.tag in (w('pPr'), w('tcPr'), w('sectPr')):
                node[:] = sorted(node, key=lambda n: (n.tag, n.get(w('type'), '')))
            if node.tag in (w('ins'), w('del')):
                existing = next((n for n in original.root(name).iter() if n.tag == node.tag
                                 and n.get(w('author')) == node.get(w('author'))
                                 and revision_text(n) == revision_text(node)), None) if name in original.parts else None
                if existing is None:
                    node.attrib.pop(w('date'), None)
                node.attrib.pop(w('id'), None)
            if node.tag == w('delText'):
                node.tag = w('t')
        for parent in root.iter():
            previous = None
            for node in list(parent):
                if node.tag == w('proofErr'):
                    continue
                if previous is not None and node.tag in (w('ins'), w('del')) and node.tag == previous.tag and node.attrib == previous.attrib:
                    previous.extend(list(node))
                    parent.remove(node)
                else:
                    previous = node
        if root.tag == q('w15', 'commentsEx'):
            root[:] = sorted((n for n in root if set(n.attrib) != {q('w15', 'paraId')} or len(n)),
                             key=lambda n: n.get(q('w15', 'paraId'), ''))
    old_styles = {n.get(w('styleId')) for n in original.root('word/styles.xml').findall('w:style', NS)}
    styles_root = package.root('word/styles.xml')
    definitions = {n.get(w('styleId')): n for n in styles_root.findall('w:style', NS)}
    pending = list(support_styles - old_styles)
    while pending:
        identity = pending.pop()
        definition = definitions.get(identity)
        if definition is not None:
            for node in definition:
                value = node.get(w('val'))
                if node.tag in (w('basedOn'), w('link'), w('next')) and value not in old_styles | support_styles:
                    support_styles.add(value)
                    pending.append(value)
    used = {n.get(w('val')) for name in package.namelist() if name.startswith('word/') and name.endswith('.xml')
            and name != 'word/styles.xml' for n in package.root(name).iter() if n.tag in (w('pStyle'), w('rStyle'))}
    for identity in support_styles - old_styles - used:
        definition = definitions.get(identity)
        if definition is not None:
            require(definition.get(w('default'), '0') not in ('1', 'true', 'on'), 'new support style cannot change defaults')
            styles_root.remove(definition)


def sqref_state(value):
    columns = {}
    for rectangle in value.replace('$', '').split():
        if ':' not in rectangle:
            rectangle += ':' + rectangle
        c1, r1, c2, r2 = range_coordinates(rectangle)
        require(1 <= c1 <= c2 <= 16384 and 1 <= r1 <= r2 <= 1048576, 'invalid conditional format range')
        for column in range(c1, c2 + 1):
            columns.setdefault(column, []).append((r1, r2))
    for column, spans in columns.items():
        merged = []
        for low, high in sorted(spans):
            if merged and low <= merged[-1][1] + 1:
                merged[-1] = (merged[-1][0], max(merged[-1][1], high))
            else:
                merged.append((low, high))
        columns[column] = tuple(merged)
    return tuple(sorted(columns.items()))


def relative_formula(formula, origin):
    column, row = re.fullmatch(r'([A-Za-z]+)([1-9][0-9]*)', origin).groups()
    column, row = column_number(column.upper()), int(row)
    def reference(match):
        absolute_col, col, absolute_row, number = match.groups()
        col, number = column_number(col.upper()), int(number)
        if col > 16384 or number > 1048576:
            return match[0]
        return '\0' + json.dumps([['absolute' if absolute_col else 'relative', col if absolute_col else col - column],
                                  ['absolute' if absolute_row else 'relative', number if absolute_row else number - row]]) + '\0'
    pieces = re.split(r'("(?:[^"]|"")*"|\'(?:[^\']|\'\')*\'|\[[^\]]*\])', formula)
    for index in range(0, len(pieces), 2):
        pieces[index] = re.sub(r'(?<![\w.])(?!(?:[A-Z]+[0-9]+)\s*\()(\$?)([A-Z]{1,3})(\$?)([0-9]+)(?![\w.!])',
                               reference, pieces[index], flags=re.IGNORECASE)
    return json.dumps(''.join(pieces))


def normalize_complex_xlsx(package, spec):
    styles_root = package.root(related(package, 'xl/workbook.xml', 'styles'))
    dxfs = styles_root.find('s:dxfs', NS)
    pool = list(dxfs) if dxfs is not None else []
    for name in package.namelist():
        if not name.startswith('xl/') or not name.endswith('.xml'):
            continue
        for node in package.root(name).iter():
            if 'dxfId' in node.attrib:
                dxf = copy.deepcopy(indexed(pool, node.get('dxfId'), 'differential format'))
                for font in dxf.findall('s:font', NS):
                    for flag in font:
                        if local(flag.tag) in ('b', 'i'):
                            flag.set('val', '1' if flag.get('val', '1') in ('1', 'true') else '0')
                    font[:] = sorted(font, key=lambda n: n.tag)
                dxf[:] = sorted(dxf, key=lambda n: n.tag)
                for color in dxf.iter():
                    if color.get('rgb') is not None:
                        color.set('rgb', color.get('rgb').upper())
                for pattern in dxf.findall('s:fill/s:patternFill', NS):
                    if pattern.get('patternType') == 'solid':
                        for n in pattern.findall('s:bgColor', NS):
                            pattern.remove(n)
                node.set('dxfId', json.dumps(canonical(dxf)))
    if dxfs is not None:
        styles_root.remove(dxfs)
    for change in spec.get('autofilter_sort', []):
        root = package.root(sheet_map(package)[change['sheet']])
        auto = root.find('s:autoFilter', NS)
        for parent in (root, auto):
            if parent is None:
                continue
            for state in parent.findall('s:sortState', NS):
                require(range_coordinates(state.get('ref').replace('$', '')) == range_coordinates(change['range']),
                        'sortState range differs')
                conditions = state.findall('s:sortCondition', NS)
                require(len(conditions) == len(change['keys']), 'sortState key count differs')
                _, low, _, high = range_coordinates(change['range'])
                for condition, key in zip(conditions, change['keys']):
                    compare(f'{key["column"]}{low}:{key["column"]}{high}', condition.get('ref', '').replace('$', ''), 'sort key')
                    compare(not key['ascending'], condition.get('descending', '0') in ('1', 'true'), 'sort direction')
                parent.remove(state)
    for part in sheet_map(package).values():
        root = package.root(part)
        groups = root.findall('s:conditionalFormatting', NS)
        rules = []
        for group in groups:
            sqref = group.get('sqref')
            origin = sqref.split()[0].split(':')[0].replace('$', '')
            for rule in group:
                for formula in rule.findall('s:formula', NS):
                    formula.text = relative_formula(formula.text or '', origin)
                rule.set('complexSqref', sqref)
                if rule.get('stopIfTrue', '0') in ('0', 'false'):
                    rule.attrib.pop('stopIfTrue', None)
                rules.append(rule)
            root.remove(group)
        rules.sort(key=lambda n: int(n.get('priority')))
        merged = []
        for rule in rules:
            sqref = rule.attrib.pop('complexSqref')
            rule.attrib.pop('priority')
            state = canonical(rule)
            if merged and canonical(merged[-1][0]) == state:
                merged[-1] = (merged[-1][0], merged[-1][1] + ' ' + sqref)
            else:
                merged.append((rule, sqref))
        for index, (rule, sqref) in enumerate(merged, 1):
            rule.set('priority', str(index))
            group = ET.SubElement(root, q('s', 'conditionalFormatting'), sqref=json.dumps(sqref_state(sqref)))
            group.append(rule)
    for change in spec.get('chart', []):
        root = package.root(change['part'])
        chart = root.find('c:chart', NS)
        title = chart.find('c:title/c:tx', NS)
        require(title is not None, 'missing chart title')
        require(title.find('c:rich', NS) is not None or title.find('c:strRef/c:f', NS) is not None,
                'chart title must use rich text or a string reference')
        title_text = ''.join(n.text or '' for n in title.iter(q('a', 't')))
        reference = title.find('c:strRef/c:f', NS)
        if reference is not None:
            title_text = chart_reference_text(package, reference.text)
        if not title_text:
            title_text = ''.join(n.text or '' for n in title.iter(q('c', 'v')))
        rich = title.find('c:rich', NS)
        formatting = {'body': None, 'list': None, 'paragraphs': [(None, None, [None])]}
        if rich is not None:
            formatting = {'body': properties(rich.find('a:bodyPr', NS)), 'list': properties(rich.find('a:lstStyle', NS)),
                          'paragraphs': [(properties(p.find('a:pPr', NS)), properties(p.find('a:endParaRPr', NS)),
                                          sorted(set(properties(r.find('a:rPr', NS)) for r in p.findall('a:r', NS)), key=repr))
                                         for p in rich.findall('a:p', NS)]}
        title.set('complexTextFormatting', json.dumps(formatting))
        title[:] = []
        title.text = title_text
        for node in chart.findall('.//c:ser', NS):
            tx = node.find('c:tx', NS)
            name = tx.findtext('c:v', namespaces=NS)
            if name is None:
                reference = tx.find('c:strRef/c:f', NS)
                require(reference is not None, 'chart series name missing')
                name = chart_reference_text(package, reference.text)
            tx[:] = []
            tx.text = name
            for field in ('cat', 'val', 'xVal', 'yVal'):
                target = node.find('c:' + field, NS)
                if target is None:
                    continue
                if field in ('xVal', 'yVal', 'val'):
                    require(target.find('c:numRef', NS) is not None, 'numeric chart data must use numRef')
                formula = unique(target.findall('.//c:f', NS), 'chart data must have one range reference').text
                target[:] = []
                target.text = chart_range_state(formula)


def chart_range_state(formula):
    match = re.fullmatch(r"('(?:[^']|'')+'|[^!]+)!(\$?[A-Za-z]+\$?[1-9][0-9]*)(?::(\$?[A-Za-z]+\$?[1-9][0-9]*))?", formula or '')
    require(match is not None, 'chart data must reference a workbook cell range')
    sheet, first, last = match.groups()
    sheet = sheet[1:-1].replace("''", "'") if sheet.startswith("'") else sheet
    return json.dumps((sheet.casefold(), first.replace('$', '').upper(), (last or first).replace('$', '').upper()))


def chart_reference_text(package, formula):
    match = re.fullmatch(r"('(?:[^']|'')+'|[^!]+)!\$?([A-Z]+)\$?([0-9]+)", formula or '')
    require(match is not None, 'chart text must reference one workbook cell')
    sheet, column, row = match.groups()
    sheet = sheet[1:-1].replace("''", "'") if sheet.startswith("'") else sheet
    model = unique((n for n in workbook(package, real=True)['sheets'] if n['name'] == sheet), 'chart name sheet missing')
    value = model['cells'].get(column + row, {}).get('value')
    require(isinstance(value, str), 'chart name reference must contain text')
    return value


def normalize_complex_slide_comments(package, spec, original):
    if not spec.get('comments'):
        return
    authors_part = related(package, 'ppt/presentation.xml', 'commentAuthors')
    authors = package.root(authors_part)
    original_part = related(original, 'ppt/presentation.xml', 'commentAuthors')
    old_names = {n.get('name') for n in original.root(original_part)}
    ordered = sorted(authors, key=lambda n: (n.get('name'), n.get('id')))
    author_map = {n.get('id'): str(index) for index, n in enumerate(ordered)}
    new_bodies = {(change['author'], change['text']) for change in spec['comments']}
    names = {n.get('id'): n.get('name') for n in authors}
    for part in package.namelist():
        if not part.endswith('.xml') or package.root(part).tag != q('p', 'cmLst'):
            continue
        root = package.root(part)
        comments = sorted(root, key=lambda n: (names[n.get('authorId')], n.findtext('p:text', default='', namespaces=NS), n.get('idx')))
        identities = {(n.get('authorId'), n.get('idx')): str(index) for index, n in enumerate(comments)}
        for node in comments:
            author_id, index = node.get('authorId'), node.get('idx')
            thread = node.find('.//p15:threadingInfo', NS)
            parent_comment = thread.find('p15:parentCm', NS) if thread is not None else None
            if parent_comment is not None:
                parent_author, parent_index = parent_comment.get('authorId'), parent_comment.get('idx')
                require((parent_author, parent_index) in identities, 'slide reply parent missing')
                parent_comment.set('authorId', author_map[parent_author])
                parent_comment.set('idx', identities[(parent_author, parent_index)])
            if (names[author_id], node.findtext('p:text', namespaces=NS)) in new_bodies:
                require(parent_comment is not None, 'requested slide comment is not a reply')
                parent = copy.deepcopy(thread)
                parent.attrib.pop('timeZoneBias', None)
                text = node.findtext('p:text', namespaces=NS)
                node.attrib = {'authorId': author_id, 'idx': index}
                node[:] = []
                ET.SubElement(node, q('p', 'text')).text = text
                node.append(parent)
            node.set('authorId', author_map[author_id])
            node.set('idx', identities[(author_id, index)])
        root[:] = comments
    for node in ordered:
        node.set('id', author_map[node.get('id')])
        node.attrib.pop('lastIdx', None)
        if node.get('name') not in old_names:
            for key in ('initials', 'clrIdx'):
                node.attrib.pop(key, None)
    authors[:] = ordered


def normalize_complex_table_ids(package, original):
    for part in slide_parts(package):
        old = {value for node in original.root(part).iter() if local(node.tag) in ('rowId', 'colId', 'creationId')
               for value in node.attrib.values()} if part in original.parts else set()
        for table in package.root(part).findall('.//a:tbl', NS):
            seen = set()
            counter = 0
            for node in table.iter():
                if local(node.tag) not in ('rowId', 'colId', 'creationId'):
                    continue
                for key, value in list(node.attrib.items()):
                    if local(key) not in ('val', 'id'):
                        continue
                    identity = (node.tag, key, value)
                    require(identity not in seen, 'duplicate table identity')
                    seen.add(identity)
                    if value not in old:
                        node.set(key, f'<new table identity {counter}>')
                        counter += 1


def complex_slide_metadata(original, package, spec):
    if not spec.get('duplicate_slide') or 'docProps/app.xml' not in package.parts:
        return
    source = original.root('docProps/app.xml')
    headings = next((n for n in source if local(n.tag) == 'HeadingPairs'), None)
    titles = next((n for n in source if local(n.tag) == 'TitlesOfParts'), None)
    if headings is None or titles is None:
        return
    updated_headings, updated_titles = copy.deepcopy(headings), copy.deepcopy(titles)
    variants, vector = list(updated_headings[0]), updated_titles[0]
    offset = 0
    for index in range(0, len(variants), 2):
        count = variants[index + 1][0]
        size = int(count.text)
        if variants[index][0].text == 'Slide Titles':
            if size == 0:
                return
            items = list(vector)
            group = items[offset:offset + size]
            for change in spec['duplicate_slide']:
                group.insert(change['after'], copy.deepcopy(group[change['slide'] - 1]))
            for title, part in zip(group, slide_parts(package)):
                shapes = [n for n in package.root(part).findall('p:cSld/p:spTree/p:sp', NS)
                          if n.find('p:nvSpPr/p:nvPr/p:ph', NS) is not None
                          and n.find('p:nvSpPr/p:nvPr/p:ph', NS).get('type') in ('title', 'ctrTitle')]
                if len(shapes) == 1:
                    title.text = node_text(shapes[0], 'a')
            vector[:] = items[:offset] + group + items[offset + size:]
            vector.set('size', str(len(vector)))
            count.text = str(len(group))
            break
        offset += size
    else:
        return
    root = package.root('docProps/app.xml')
    found = [next((n for n in root if n.tag == item.tag), None) for item in (headings, titles)]
    state = tuple(canonical(n) for n in found)
    require(state in (tuple(map(canonical, (headings, titles))), tuple(map(canonical, (updated_headings, updated_titles)))),
            'unexpected duplicated slide title metadata')
    for before, after in zip((headings, titles), found):
        index = list(root).index(after)
        root.remove(after)
        root.insert(index, copy.deepcopy(before))


def accept_filter_database(expected, actual, spec):
    for change in spec.get('autofilter_sort', []):
        found = []
        for package in (expected, actual):
            root = package.root('xl/workbook.xml')
            sheets = [n.get('name') for n in root.findall('s:sheets/s:sheet', NS)]
            index = str(sheets.index(change['sheet'])) if change['sheet'] in sheets else None
            found.append([(root.find('s:definedNames', NS), n) for n in root.findall('s:definedNames/s:definedName', NS)
                          if n.get('name', '').casefold() == '_xlnm._filterdatabase' and n.get('localSheetId') == index])
        if index is None or found[0] or len(found[1]) != 1:
            continue
        names, item = found[1][0]
        sheet, _, ref = (item.text or '').rpartition('!')
        sheet = sheet[1:-1].replace("''", "'") if len(sheet) > 1 and sheet[0] == sheet[-1] == "'" else sheet
        if (sheet == change['sheet'] and ref == re.sub(r'([A-Z]+)(\d+)', r'$\1$\2', change['ref'])
                and item.get('hidden') in ('1', 'true') and set(item.attrib) == {'name', 'localSheetId', 'hidden'}):
            names.remove(item)
            if not len(names) and expected.root('xl/workbook.xml').find('s:definedNames', NS) is None:
                actual.root('xl/workbook.xml').remove(names)


def accept_full_calc_on_load(expected, actual):
    calculation = expected.root('xl/workbook.xml').find('s:calcPr', NS)
    saved_calculation = actual.root('xl/workbook.xml').find('s:calcPr', NS)
    if calculation is not None and saved_calculation is not None and saved_calculation.get('fullCalcOnLoad') in ('1', 'true'):
        calculation.set('fullCalcOnLoad', saved_calculation.get('fullCalcOnLoad'))


def check_complex(task, source, revised):
    original, expected, actual = RealPackage(source), RealPackage(source), RealPackage(revised)
    spec, fmt = task['expected'], task['format']
    for package in (original, expected, actual):
        for name in package.namelist():
            if name.startswith('[trash]/'):
                del package.parts[name]
        types = package.root('[Content_Types].xml')
        types[:] = [n for n in types if not n.get('PartName', '').lstrip('/').startswith('[trash]/')]
    complex_integrity(actual, original)
    complex_expected_package(expected, spec, fmt)
    complex_xml_order(expected)
    complex_xml_order(actual)
    if fmt == 'docx':
        normalize_comment_metadata(expected, original)
        normalize_comment_metadata(actual, original)
    align_complex_parts(expected, actual, original, fmt)
    if fmt == 'xlsx':
        accept_filter_database(expected, actual, spec)
        normalize_complex_xlsx(expected, spec)
        normalize_complex_xlsx(actual, spec)
        complex_xml_order(expected)
        complex_xml_order(actual)
        real_metadata(expected)
        real_metadata(actual)
        accept_full_calc_on_load(expected, actual)
        compare(workbook(expected, real=True), workbook(actual, real=True))
    else:
        if fmt == 'docx':
            normalize_complex_word(expected, spec, original)
            normalize_complex_word(actual, spec, original)
        else:
            complex_slide_metadata(original, expected, spec)
            complex_slide_metadata(original, actual, spec)
            normalize_complex_slide_comments(expected, spec, original)
            normalize_complex_slide_comments(actual, spec, original)
            normalize_complex_table_ids(expected, original)
            normalize_complex_table_ids(actual, original)
        ids = {value for _, value in real_ids(original)} if fmt == 'docx' else set()
        compare(real_package_state(expected, True, ids), real_package_state(actual))


def check(task, output):
    output = Path(output)
    if task.get('tier') == 'complex' and 'answer' in task['expected']:
        require(set(task['expected']) == {'answer'}, 'answer cannot be combined with document edits')
        answer = output if output.name == 'answer.json' else output.parent / 'answer.json'
        require(answer.is_file(), 'missing answer.json')
        def pairs(items):
            result = {}
            for key, value in items:
                require(key not in result, 'duplicate answer JSON key')
                result[key] = value
            return result
        value = json.loads(answer.read_text(), object_pairs_hook=pairs,
                           parse_constant=lambda value: require(False, 'non-finite answer JSON'))
        compare(task['expected']['answer'], value, 'answer')
        return 'exact answer verified'
    require(output.is_file(), f'missing output file: {output.name}')
    with ZipFile(BytesIO(prepare_input(task, fixture(task).read_bytes()))) as source, ZipFile(output) as revised:
        validate_package(revised, ignore_trash=task.get('tier') == 'complex')
        if task.get('tier') == 'complex':
            require(task.get('suite') == 'v3', 'complex tier requires v3')
            check_complex(task, source, revised)
            return 'complex edits, package integrity and preserved content verified'
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
    task = next((task for task in tasks(args.suite, 'all' if args.suite == 'v3' else 'basic') if task['id'] == args.task), None)
    if task is None:
        parser.error(f'unknown {args.suite} task: {args.task}')
    fixture(task)
    try:
        print(f'PASS: {task["id"]}: {check(task, args.output)}')
    except Exception as error:
        print(f'FAIL: {task["id"]}: {error}')
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
