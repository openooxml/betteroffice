"""Build synthetic table-width fixtures for the Microsoft Word oracle."""
import sys
import zipfile
from pathlib import Path

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"


def paragraph(text, properties=""):
    return (f'<w:p><w:pPr>{properties}</w:pPr><w:r>'
            f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p>')


def cell(text, width=0, extra="", properties=""):
    kind = "dxa" if width else "auto"
    return (f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="{kind}"/>'
            f'{extra}</w:tcPr>{paragraph(text, properties)}</w:tc>')


def build(directory, name, mode, grid, widths, extra="", cell_extra="", indent="", table_width=0):
    phrase = "alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu"
    rows = []
    for index in range(1, 81):
        rows.append('<w:tr>'
                    + cell(f'ID {index:03}', widths[0], cell_extra)
                    + cell(f'Row {index:03} {phrase}', widths[1], properties=indent)
                    + '</w:tr>')
    layout = f'<w:tblLayout w:type="{mode}"/>' if mode else ""
    kind = "dxa" if table_width else "auto"
    body = (f'<w:tbl><w:tblPr><w:tblW w:w="{table_width}" w:type="{kind}"/>'
            f'{layout}{extra}<w:tblCellMar><w:top w:w="0" w:type="dxa"/>'
            '<w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>'
            '<w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>'
            + ''.join(f'<w:gridCol w:w="{width}"/>' for width in grid)
            + '</w:tblGrid>' + ''.join(rows) + '</w:tbl>' + paragraph('End'))
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    styles = (f'<w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
              '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" '
              'w:line="240" w:lineRule="exact"/><w:widowControl w:val="0"/>'
              '</w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/></w:style></w:styles>')
    parts = {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + ''.join(f'<Override PartName="/word/{part}.xml" ContentType="{CT}.{kind}+xml"/>'
                  for part, kind in [('document', 'document.main'), ('styles', 'styles'), ('settings', 'settings')])
        + '</Types>',
        '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rDoc" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>',
        'word/_rels/document.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rStyles" Type="{REL}/styles" Target="styles.xml"/>'
        f'<Relationship Id="rSettings" Type="{REL}/settings" Target="settings.xml"/></Relationships>',
        'word/document.xml': f'<w:document {W}><w:body>{body}{sect}</w:body></w:document>',
        'word/styles.xml': styles,
        'word/settings.xml': f'<w:settings {W}><w:compat><w:compatSetting w:name="compatibilityMode" '
        'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>',
    }
    directory.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(directory / f'{name}.docx', 'w') as archive:
        for key, value in parts.items():
            info = zipfile.ZipInfo(key, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + value)


if __name__ == '__main__':
    directory = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent
    build(directory, 'fixed-preferred', 'fixed', [1440, 1440], [1440, 5760], table_width=7200)
    build(directory, 'fixed-span', 'fixed', [720, 720, 1440], [1440, 5760], cell_extra='<w:gridSpan w:val="2"/>', table_width=7200)
    build(directory, 'autofit', 'autofit', [4320, 1440], [0, 0])
    build(directory, 'autofit-span', 'autofit', [2160, 2160, 1440], [0, 0], cell_extra='<w:gridSpan w:val="2"/>')
    build(directory, 'default-autofit', '', [4320, 1440], [0, 0])
    build(directory, 'nowrap', 'autofit', [720, 720], [0, 0], cell_extra='<w:noWrap/>')
    build(directory, 'margins-indent', 'autofit', [4320, 1440], [0, 0],
          extra='<w:tblInd w:w="720" w:type="dxa"/>',
          cell_extra='<w:tcMar><w:left w:w="180" w:type="dxa"/><w:right w:w="180" w:type="dxa"/></w:tcMar>',
          indent='<w:ind w:left="180" w:right="180" w:firstLine="180"/>')
