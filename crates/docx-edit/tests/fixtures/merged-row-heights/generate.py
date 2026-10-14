#!/usr/bin/env python3
"""Generate synthetic merged-cell row-height fixtures."""

import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"
WIDTHS = (1800, 1800, 2160)


def paragraph(text):
    return (
        '<w:p><w:pPr><w:keepNext w:val="0"/><w:keepLines w:val="0"/>'
        '<w:widowControl w:val="0"/>'
        '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>'
        '</w:pPr><w:r><w:t xml:space="preserve">'
        f'{escape(text)}</w:t></w:r></w:p>'
    )


def cell(column, text=(), merge=None):
    merged = '' if merge is None else f'<w:vMerge w:val="{merge}"/>'
    content = ''.join(paragraph(line) for line in text) or paragraph('')
    return (
        f'<w:tc><w:tcPr><w:tcW w:w="{WIDTHS[column]}" w:type="dxa"/>'
        f'{merged}<w:vAlign w:val="top"/></w:tcPr>{content}</w:tc>'
    )


def table(repeated, reverse):
    requirements = [('Short', 8), ('Tall', 10)]
    if reverse:
        requirements.reverse()
    rows = []
    for index in range(3):
        cells = []
        for column, (label, lines) in enumerate(requirements):
            text = [f'{label} {line + 1:02d}' for line in range(lines)] if index == 0 else []
            cells.append(cell(column, text, 'restart' if index == 0 else 'continue'))
        cells.append(cell(2, [f'H{index + 1:03d}']))
        header = '<w:tblHeader/>' if repeated else ''
        rows.append(f'<w:tr><w:trPr><w:cantSplit/>{header}</w:trPr>{"".join(cells)}</w:tr>')
    for index in range(60 if repeated else 20):
        cells = cell(0, ['Left']) + cell(1, ['Right']) + cell(2, [f'B{index + 1:03d}'])
        rows.append(f'<w:tr><w:trPr><w:cantSplit/></w:trPr>{cells}</w:tr>')
    margins = ''.join(f'<w:{side} w:w="0" w:type="dxa"/>' for side in ('top', 'left', 'bottom', 'right'))
    borders = ''.join(f'<w:{side} w:val="nil"/>' for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    grid = ''.join(f'<w:gridCol w:w="{width}"/>' for width in WIDTHS)
    return (
        '<w:tbl><w:tblPr><w:tblW w:w="5760" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        f'<w:tblBorders>{borders}</w:tblBorders><w:tblCellMar>{margins}</w:tblCellMar>'
        f'</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{"".join(rows)}</w:tbl>'
    )


def document(path, repeated, reverse):
    section = (
        '<w:sectPr><w:pgSz w:w="7200" w:h="9000"/>'
        '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" '
        'w:header="300" w:footer="300" w:gutter="0"/></w:sectPr>'
    )
    styles = (
        f'<w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
        '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" '
        'w:line="240" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/></w:style></w:styles>'
    )
    parts = {
        '[Content_Types].xml': (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/word/document.xml" ContentType="{CT}.document.main+xml"/>'
            f'<Override PartName="/word/styles.xml" ContentType="{CT}.styles+xml"/>'
            f'<Override PartName="/word/settings.xml" ContentType="{CT}.settings+xml"/></Types>'
        ),
        '_rels/.rels': (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rDoc" Type="{REL}/officeDocument" Target="word/document.xml"/>'
            '</Relationships>'
        ),
        'word/_rels/document.xml.rels': (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rStyles" Type="{REL}/styles" Target="styles.xml"/>'
            f'<Relationship Id="rSettings" Type="{REL}/settings" Target="settings.xml"/>'
            '</Relationships>'
        ),
        'word/document.xml': (
            f'<w:document {W}><w:body>{table(repeated, reverse)}'
            f'{paragraph("END")}{section}</w:body></w:document>'
        ),
        'word/styles.xml': styles,
        'word/settings.xml': (
            f'<w:settings {W}><w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
        ),
    }
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, xml in parts.items():
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + xml)
    print(path)


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent
    out.mkdir(parents=True, exist_ok=True)
    for repeated, stem in [(False, 'same-span'), (True, 'repeated-header')]:
        for reverse, suffix in [(False, ''), (True, '-reversed')]:
            document(out / f'{stem}{suffix}.docx', repeated, reverse)


if __name__ == '__main__':
    main()
