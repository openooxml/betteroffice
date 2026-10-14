#!/usr/bin/env python3
"""Fixtures: a table row with a minimum height at the page bottom, by room left and content.
usage: generate.py <outdir>"""
import struct, zlib, zipfile, sys
from pathlib import Path

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"

def png():
    raw = b"\x00" + b"\x80\x80\x80"  # 1x1 grey
    def chunk(t, d): return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xffffffff)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")

def image(pt_w, pt_h, n=1):
    cx, cy = int(pt_w * 12700), int(pt_h * 12700)
    return (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{n}" name="p{n}"/>'
            f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="{n}" name="p{n}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="rImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')

def p(text, ppr="", extra=""):
    t = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f'<w:p><w:pPr>{ppr}</w:pPr>{t}{extra}</w:p>'

def build(path, body, normal_spacing='<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>', table_style=""):
    styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>'
              f'<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr>{normal_spacing}</w:pPr></w:style>'
              f'<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
              f'{table_style}</w:styles>')
    sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
    parts = {
        "[Content_Types].xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="{CT}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="{CT}.styles+xml"/><Override PartName="/word/settings.xml" ContentType="{CT}.settings+xml"/></Types>',
        "_rels/.rels": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>',
        "word/_rels/document.xml.rels": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rStyles" Type="{REL}/styles" Target="styles.xml"/><Relationship Id="rSettings" Type="{REL}/settings" Target="settings.xml"/><Relationship Id="rImg" Type="{REL}/image" Target="media/image1.png"/></Relationships>',
        "word/document.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>',
        "word/styles.xml": styles,
        "word/settings.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>',
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for k, v in parts.items():
            info = zipfile.ZipInfo(k, date_time=(1980, 1, 1, 0, 0, 0)); info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, v)
        info = zipfile.ZipInfo("word/media/image1.png", date_time=(1980, 1, 1, 0, 0, 0))
        z.writestr(info, png())

def table(cells_xml, style=None):
    sty = f'<w:tblStyle w:val="{style}"/>' if style else ''
    rows = ''.join(f'<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>{c}</w:tc></w:tr>' for c in cells_xml)
    return f'<w:tbl><w:tblPr>{sty}<w:tblW w:w="9000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>{rows}</w:tbl>'

B = "<w:tblBorders>" + "".join(f'<w:{n} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for n in ["top", "left", "bottom", "right", "insideH", "insideV"]) + "</w:tblBorders>"
exact12 = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>'
tpl = '<w:spacing w:line="276" w:lineRule="auto"/>'

def row(label, lines, height):
    tr = f'<w:trPr><w:trHeight w:val="{height}"/></w:trPr>' if height else ''
    content = "".join(p(f"{label} line {i + 1}", tpl) for i in range(lines))
    return f'<w:tr>{tr}<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>{content}</w:tc></w:tr>'

def doc(name, filler, lines, height):
    body = "".join(p(f"Filler {i + 1:02d}", exact12) for i in range(filler))
    table = (f'<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>{B}<w:tblLayout w:type="fixed"/>'
             f'<w:tblCellMar><w:left w:w="57" w:type="dxa"/><w:right w:w="57" w:type="dxa"/></w:tblCellMar></w:tblPr>'
             f'<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>{row("Test", lines, height)}{row("After", 1, None)}</w:tbl>')
    build(out / f"{name}.docx", body + table + p("Bottom"), exact12)

out = Path(sys.argv[1]); out.mkdir(parents=True, exist_ok=True)
for name, filler, lines, height in [
    ("split-min850-2l-room38", 55, 2, 850),
    ("split-min850-2l-room26", 56, 2, 850),
    ("split-min1700-2l-room50", 54, 2, 1700),
    ("split-min850-6l-room50", 54, 6, 850),
    ("split-nomin-6l-room50", 54, 6, None),
    ("split-min1700-4l-room50", 54, 4, 1700),
]:
    doc(name, filler, lines, height)
print("ok")
