import hashlib
import json
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile


W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def accepted_text(element):
    if element.tag == W + 'del':
        return ''
    if element.tag == W + 't':
        return element.text or ''
    return ''.join(accepted_text(child) for child in element)


def paragraphs(xml):
    return [accepted_text(p) for p in ET.fromstring(xml).iter(W + 'p')]


def check(root, task, replacements):
    source = root / f'{task}.docx'
    output = root / f'{task}-updated.docx'
    with ZipFile(source) as before, ZipFile(output) as after:
        for part in ['word/document.xml', 'word/header1.xml']:
            expected = paragraphs(before.read(part))
            for old, new in replacements:
                expected = [text.replace(old, new, 1) for text in expected]
            assert paragraphs(after.read(part)) == expected, f'{task}: content mismatch in {part}'
        for part in ['word/styles.xml', 'customXml/preserved.xml']:
            assert before.read(part) == after.read(part), f'{task}: changed untouched part {part}'
        if task == 'revenue':
            runs = ET.fromstring(after.read('word/document.xml')).iter(W + 'r')
            revenue = next(run for run in runs if accepted_text(run) == '€5.1 million')
            assert revenue.find(f'{W}rPr/{W}b') is not None, 'revenue lost bold'
            assert revenue.find(f'{W}rPr/{W}color').get(W + 'val') == '0066AA', 'revenue lost color'
        if task == 'cross-story':
            revisions = []
            for part in ['word/document.xml', 'word/header1.xml']:
                tree = ET.fromstring(after.read(part))
                revisions.extend(list(tree.iter(W + 'ins')) + list(tree.iter(W + 'del')))
            assert len(revisions) == 4, 'expected a tracked replacement in body and header'
            assert all(r.get(W + 'author') == 'Luna Review' for r in revisions), 'wrong revision author'
            assert all(r.get(W + 'date') == '2026-09-27T12:00:00Z' for r in revisions), 'wrong revision date'
    print(f'PASS {task}: exact expected content, preserved parts, and applicable formatting/revisions')


root = Path(sys.argv[1])
hashes = root / 'source-hashes.json'
if hashes.exists():
    for name, expected in json.loads(hashes.read_text()).items():
        assert hashlib.sha256((root / name).read_bytes()).hexdigest() == expected, f'source changed: {name}'
else:
    print('Source-byte check skipped: no source-hashes.json found.')
tasks = {
    'revenue': [('€4.2 million', '€5.1 million'), ('15 October 2026', '12 November 2026')],
    'wording': [('The forecast is draft; the appendix is draft.', 'The forecast is draft; the appendix is approved.'), ('Table target: approved', 'Table target: pending review')],
    'cross-story': [('Executive summary', 'Board summary'), ('Internal — Project Aurora', 'Board review — Project Aurora')],
}
for task in sys.argv[2:] or tasks:
    check(root, task, tasks[task])
