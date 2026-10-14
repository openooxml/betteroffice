"""Check the format prototype evaluation exports."""
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

root = Path(sys.argv[1])
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}


def cells(path):
    with ZipFile(path) as archive:
        sheet = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
        return {cell.attrib['r']: cell for cell in sheet.findall('.//s:c', ns)}


def value(cell):
    return cell.findtext('s:v', namespaces=ns)


def slide(path):
    with ZipFile(path) as archive:
        return ET.fromstring(archive.read('ppt/slides/slide1.xml'))


original = cells(root / 'xlsx/budget.xlsx')
revised = cells(root / 'xlsx/budget-revised.xlsx')
assert value(original['B3']) == '100'
assert value(revised['B3']) == '1000'
assert revised['E3'].findtext('s:f', namespaces=ns) == 'D3*2'
assert value(revised['D3']) == '1057'
assert value(revised['E3']) == '2114'
for address in original.keys() & revised.keys():
    assert original[address].get('s') == revised[address].get('s'), address
    if address not in ('B3', 'D3', 'E3'):
        assert ET.tostring(original[address]) == ET.tostring(revised[address]), address

for task, expected in [('pptx', 'Revenue €5.1 million. Risk opportunity. 😀'),
                       ('mixed', 'Revenue €6.0 million. Risk risk. 😀')]:
    source = root / task / 'slides.pptx'
    exported = root / task / 'slides-revised.pptx'
    before = slide(source)
    after = slide(exported)
    assert ''.join(before.itertext()) == 'Revenue €4.2 million. Risk risk. 😀Keep this paragraph.'
    paragraphs = [''.join(p.itertext()) for p in after.findall('.//a:p', ns)]
    assert paragraphs == [expected, 'Keep this paragraph.'], paragraphs
    assert [ET.tostring(p) for p in before.findall('.//a:rPr', ns)] == [ET.tostring(p) for p in after.findall('.//a:rPr', ns)]
    with ZipFile(source) as a, ZipFile(exported) as b:
        for name in a.namelist():
            if not name.endswith('/') and name != 'ppt/slides/slide1.xml':
                assert a.read(name) == b.read(name), name

mixed = cells(root / 'mixed/budget-revised.xlsx')
assert value(mixed['B3']) == '700'
assert value(mixed['D3']) == '757'
print('PASS: formula results, exact slide changes, source contents, and preserved formatting/package parts')
