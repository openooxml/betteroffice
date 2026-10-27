#!/usr/bin/env python3
"""Deterministic synthetic long DOCX; seed 23024; CC0 dedication."""
import argparse
import hashlib
import random
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED


def build(pages, seed):
    rng = random.Random(seed)
    words = 'river stone meadow forest valley morning light water field path cloud season garden green blue quiet clear'.split()
    w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    def para(text, props=''):
        return f'<w:p><w:pPr>{props}</w:pPr><w:r><w:t>{escape(text)}</w:t></w:r></w:p>'
    body = []
    for i in range(pages):
        body.append(para(f'Long document section {i+1}', ('<w:pageBreakBefore/>' if i else '') + '<w:pStyle w:val="Heading1"/>'))
        if i == 0:
            body.append(para('This synthetic document exercises opening scrolling and editing in long documents using body text headings lists and simple tables.'))
        for j in range(4):
            body.append(para(' '.join(rng.choice(words) for _ in range(48)).capitalize() + '.'))
        for j in range(2):
            body.append(para(f'Section {i+1} item {j+1} ' + ' '.join(rng.choice(words) for _ in range(10)), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'))
        rows = []
        for r in range(3):
            cells = ''.join(f'<w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/></w:tcPr>{para(f"Section {i+1} row {r+1} column {c+1}")}</w:tc>' for c in range(2))
            rows.append(f'<w:tr>{cells}</w:tr>')
        edges = ''.join(f'<w:{e} w:val="single" w:sz="4" w:color="D9D9D9"/>' for e in ['top','left','bottom','right','insideH','insideV'])
        body.append(f'<w:tbl><w:tblPr><w:tblW w:w="8640" w:type="dxa"/><w:tblBorders>{edges}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4320"/><w:gridCol w:w="4320"/></w:tblGrid>{"".join(rows)}</w:tbl>')
    body.append('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>')
    parts = {
      '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>',
      '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
      'word/_rels/document.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>',
      'word/document.xml': f'<w:document xmlns:w="{w}"><w:body>{"".join(body)}</w:body></w:document>',
      'word/styles.xml': f'<w:styles xmlns:w="{w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="100" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style></w:styles>',
      'word/numbering.xml': f'<w:numbering xmlns:w="{w}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
    }
    return parts


def main():
    p = argparse.ArgumentParser()
    p.add_argument('--pages', type=int, required=True)
    p.add_argument('--seed', type=int, default=23024)
    p.add_argument('--out', type=Path, required=True)
    a = p.parse_args()
    a.out.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(a.out, 'w', compression=ZIP_DEFLATED) as z:
        for name, xml in sorted(build(a.pages, a.seed).items()):
            info = ZipInfo(name, (2020, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + xml)
    print(hashlib.sha256(a.out.read_bytes()).hexdigest(), a.out, f'explicit_sections={a.pages} seed={a.seed}')

if __name__ == '__main__':
    main()
