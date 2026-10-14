#!/usr/bin/env python3
"""Fixtures: complex fields whose code spans paragraph marks, in the body and in a header.
usage: generate.py <outdir>"""
import sys
import zipfile
from pathlib import Path

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"


def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def instr(text):
    return f'<w:r><w:instrText xml:space="preserve">{text}</w:instrText></w:r>'


def fld(kind):
    return f'<w:r><w:fldChar w:fldCharType="{kind}"/></w:r>'


def para(inner):
    return f"<w:p>{inner}</w:p>"


def spanning_field(label):
    """`label` then an IF field whose code runs over three paragraph marks."""
    return (para(run(label) + fld("begin") + instr(" IF 1 = 1 "))
            + para(instr('"yes" '))
            + para(instr('"no" '))
            + para(fld("separate") + run("yes") + fld("end")))


STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'''


def document(name, header=None, body=None):
    parts = {}
    rels = [f'<Relationship Id="rStyles" Type="{REL}/styles" Target="styles.xml"/>',
            f'<Relationship Id="rSettings" Type="{REL}/settings" Target="settings.xml"/>']
    overrides = [f'<Override PartName="/word/document.xml" ContentType="{CT}.document.main+xml"/>',
                 f'<Override PartName="/word/styles.xml" ContentType="{CT}.styles+xml"/>',
                 f'<Override PartName="/word/settings.xml" ContentType="{CT}.settings+xml"/>']
    refs = ""
    if header is not None:
        parts["word/header1.xml"] = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {W}>{header}</w:hdr>'
        rels.append(f'<Relationship Id="rHeader" Type="{REL}/header" Target="header1.xml"/>')
        overrides.append(f'<Override PartName="/word/header1.xml" ContentType="{CT}.header+xml"/>')
        refs = '<w:headerReference w:type="default" r:id="rHeader"/>'
    sect = ("<w:sectPr>" + refs
            + '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    parts["word/document.xml"] = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'
    parts["word/styles.xml"] = STYLES
    parts["word/settings.xml"] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
                                  '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
    parts["word/_rels/document.xml.rels"] = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{"".join(rels)}</Relationships>'
    parts["_rels/.rels"] = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>'
    parts["[Content_Types].xml"] = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>{"".join(overrides)}</Types>'
    order = ["[Content_Types].xml", "_rels/.rels"] + sorted(k for k in parts if k.startswith("word/"))
    out = Path(sys.argv[1]) / f"{name}.docx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        for key in order:
            info = zipfile.ZipInfo(key, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, parts[key])
    print(out)


def lines(a, b):
    return "".join(para(run(f"Line {i:03d}")) for i in range(a, b + 1))


header_lines = "".join(para(run(f"Header {i}")) for i in range(1, 7))
def adjacent_field(label):
    """`label` then an IF field whose code ends at its paragraph's mark, with the separator in the next."""
    return para(run(label) + fld("begin") + instr(' IF 1 = 1 "yes" "no" ')) + para(fld("separate") + run("yes") + fld("end"))


def chained_fields(label):
    """Two spanning IF fields, the second starting in the paragraph that shows the first's result."""
    return (para(run(label) + fld("begin") + instr(" IF 1 = 1 "))
            + para(instr('"yes" "no" '))
            + para(fld("separate") + run("yes") + fld("end") + run(" and ") + fld("begin") + instr(" IF 1 = 1 "))
            + para(instr('"again" "no" '))
            + para(fld("separate") + run("again") + fld("end")))


def nested_field(label):
    """A spanning IF field whose hidden code paragraph holds a nested PAGE field."""
    return (para(run(label) + fld("begin") + instr(" IF "))
            + para(fld("begin") + instr(" PAGE ") + fld("separate") + run("1") + fld("end") + instr(' = 1 "yes" "no" '))
            + para(fld("separate") + run("yes") + fld("end")))


def cell_field():
    """A one-cell table whose cell opens with a spanning field, then a plain paragraph."""
    cell = spanning_field("Cell ") + para(run("Cell end"))
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
            + cell + '</w:tc></w:tr></w:tbl>')


document("header-field-code-paragraphs", header=header_lines + spanning_field("Header 7 "), body=lines(1, 280))
document("body-field-code-adjacent", body=lines(1, 40) + adjacent_field("Line 041 ") + lines(42, 280))
document("body-field-code-chained", body=lines(1, 40) + chained_fields("Line 041 ") + lines(42, 280))
document("body-field-code-nested", body=lines(1, 40) + nested_field("Line 041 ") + lines(42, 280))
document("cell-field-code-paragraphs", body=lines(1, 40) + cell_field() + lines(43, 280))
document("body-field-code-paragraphs", body=lines(1, 40) + spanning_field("Line 041 ") + lines(42, 280))
for name, field in [
    ("adjacent", adjacent_field),
    ("chained", chained_fields),
    ("nested", nested_field),
]:
    document(f"header-field-code-{name}", header=header_lines + field("Header 7 "), body=lines(1, 280))
document("header-field-code-cell", header=header_lines + cell_field() + para(""), body=lines(1, 280))
