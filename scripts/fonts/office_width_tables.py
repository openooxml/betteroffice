"""Read the named-family width tables in crates/ooxml-text/src/family_metrics.rs off font files.

The tables hold advance widths and hhea metrics only; the faces they are read from ship with
Office and macOS and are never committed. Point the script at the directories that hold them:

    python3 scripts/fonts/office_width_tables.py check DIR [DIR ...]
    python3 scripts/fonts/office_width_tables.py print CONST DIR [DIR ...]

`check` regenerates every table whose face it finds and reports any table that differs. Where two
directories hold the same face, the first one listed wins: list Office's own fonts first.

Japanese tables add punctuation, symbols, full-width forms and half-width katakana.
Proportional Japanese faces also include kana; ideographs keep the substitute's widths.
"""

import os
import re
import sys
from pathlib import Path

from fontTools.ttLib import TTCollection, TTFont

PLAIN_STYLE_WORDS = {'regular', 'roman', 'book', 'normal', 'bold', 'italic', 'oblique'}
JAPANESE_TABLES = {'MS_GOTHIC', 'MS_PGOTHIC', 'MS_UI_GOTHIC', 'MS_MINCHO', 'MS_PMINCHO'}
PROPORTIONAL_JAPANESE_TABLES = {'MS_PGOTHIC', 'MS_UI_GOTHIC', 'MS_PMINCHO'}
JAPANESE_EXTRA = (
    set(range(0x2000, 0x2070)) | set(range(0x20A0, 0x20D0)) |
    set(range(0x2100, 0x2150)) | set(range(0x2190, 0x219A)) |
    set(range(0x2460, 0x2474)) | set(range(0x3000, 0x3040)) |
    set(range(0xFF01, 0xFFA0)) | set(range(0xFFE0, 0xFFE7)) |
    {ord(c) for c in '∀∂∃∇∈∋∑√∞∟∠∥∧∨∩∪∫∮∴∵∽≒≠≡≦≧⊂⊃⊆⊇⊥⌒■□▲△▼▽◆◇○◎●◯★☆♀♂♪♭♯'}
)
TABLES = Path(__file__).resolve().parents[2] / 'crates/ooxml-text/src/family_metrics.rs'


def parse(source):
    tables = {}
    for name, body in re.findall(r'static ([A-Z_]+): FamilyMetrics = FamilyMetrics \{(.*?)\n\};', source, re.S):
        field = lambda key: int(re.search(key + r': (-?\d+)', body).group(1))
        ascii_widths = [int(x) for x in re.findall(r'\d+', re.search(r'ascii: \[(.*?)\]', body, re.S).group(1))]
        extra = [tuple(map(int, pair)) for pair in re.findall(r'\((\d+), (\d+)\)', body.split('extra:')[1])]
        tables[name] = (field('ascent'), field('descent'), field('line_gap'), ascii_widths, extra)
    entries = re.findall(
        r'family: "([^"]+)",\n\s+bold: (true|false),\n\s+italic: (true|false),\n\s+metrics: &([A-Z_]+),', source)
    return tables, [(family, bold == 'true', italic == 'true', name) for family, bold, italic, name in entries]


def faces(dirs):
    for root in dirs:
        for path in sorted(Path(root).rglob('*')):
            if path.suffix.lower() not in ('.ttf', '.otf', '.ttc'):
                continue
            try:
                fonts = TTCollection(path, lazy=True).fonts if path.suffix.lower() == '.ttc' else [TTFont(path, lazy=True)]
            except Exception:
                continue
            for font in fonts:
                names = font['name']
                family = (names.getDebugName(16) or names.getDebugName(1) or '').strip().lower()
                style = (names.getDebugName(17) or names.getDebugName(2) or '').lower()
                selection = font['OS/2'].fsSelection if 'OS/2' in font else 0
                bold = bool(selection & 0x20) or 'bold' in style
                italic = bool(selection & 0x01) or 'italic' in style or 'oblique' in style
                yield family, bold, italic, style, font


def table(font, codes):
    units = font['head'].unitsPerEm
    em = lambda value: round(value * 1000 / units)
    cmap = font.getBestCmap()
    metrics = font['hmtx']
    width = lambda code: em(metrics[cmap[code]][0]) if code in cmap else 0
    hhea = font['hhea']
    return (em(hhea.ascent), em(hhea.descent), em(hhea.lineGap), [width(c) for c in range(0x20, 0x7F)],
            [(c, width(c)) for c in codes if c in cmap and width(c) > 0])


def rust(name, values):
    ascent, descent, line_gap, ascii_widths, extra = values
    lines = [f'static {name}: FamilyMetrics = FamilyMetrics {{', f'    ascent: {ascent},', f'    descent: {descent},',
             f'    line_gap: {line_gap},', '    ascii: [']
    lines += ['        ' + ' '.join(f'{v},' for v in ascii_widths[i:i + 18]) for i in range(0, 95, 18)]
    lines += ['    ],', '    extra: &[']
    lines += [f'        ({code}, {value}),' for code, value in extra]
    lines += ['    ],', '};']
    return '\n'.join(lines)


def main():
    tables, entries = parse(TABLES.read_text())
    codes = {code for name, values in tables.items() if name not in JAPANESE_TABLES for code, _ in values[4]}
    mode, args = sys.argv[1], sys.argv[2:]
    if mode == 'print':
        name, dirs = args[0], args[1:]
    else:
        name, dirs = None, args
    found = {}
    for family, bold, italic, style, font in faces(dirs):
        qualifier = ' '.join(word for word in style.split() if word not in PLAIN_STYLE_WORDS)
        if qualifier and not family.endswith(qualifier):
            family = f'{family} {qualifier}'
        found.setdefault((family, bold, italic), font)
    checked = differ = 0
    for family, bold, italic, const in entries:
        if name and const != name:
            continue
        font = found.get((family, bold, italic))
        if font is None:
            continue
        extra = codes | JAPANESE_EXTRA if const in JAPANESE_TABLES else codes
        if const in PROPORTIONAL_JAPANESE_TABLES:
            extra = extra | set(range(0x3040, 0x3100))
        values = table(font, sorted(extra))
        if mode == 'print':
            print(rust(const, values))
            return
        checked += 1
        if values != tables[const]:
            differ += 1
            print(f'DIFFERS {const}')
    print(f'checked {checked} of {len(entries)} entries, {differ} differ')
    sys.exit(1 if differ else 0)


if __name__ == '__main__':
    main()
