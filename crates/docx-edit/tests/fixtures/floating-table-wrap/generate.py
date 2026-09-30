from pathlib import Path
from zipfile import ZIP_STORED, ZipFile, ZipInfo


ROOT = Path(__file__).resolve().parent
ANCHOR = "alpha beta gamma delta " * 15 + "alpha"
PIXEL = bytes.fromhex(
    "89504e470d0a1a0a0000000d4948445200000001000000010802000000907753de"
    "0000000c4944415478da63285d791b00039001faf3e482180000000049454e44ae426082"
)
FIXTURES = (
    ("right-4660", "right", 4660, False),
    ("right-4695", "right", 4695, False),
    ("right-5500", "right", 5500, False),
    ("left-5500", "left", 5500, False),
    ("right-5500-empty-anchor", "right", 5500, True),
)

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"""
PACKAGE_RELS = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="document" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
DOCUMENT_RELS = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="picture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/pixel.png"/>
</Relationships>
"""
STYLES = """<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:sz w:val="22"/><w:szCs w:val="22"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>
"""
PICTURE = """<w:p><w:r><w:drawing>
  <wp:inline distT="0" distB="0" distL="0" distR="0">
    <wp:extent cx="1524000" cy="1524000"/>
    <wp:docPr id="1" name="Synthetic pixel"/>
    <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr><pic:cNvPr id="0" name="pixel.png"/><pic:cNvPicPr/></pic:nvPicPr>
          <pic:blipFill><a:blip r:embed="picture"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
          <pic:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="1524000" cy="1524000"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing></w:r></w:p>"""
CAPTION = "<w:p><w:r><w:t>caption alpha beta</w:t><w:br/><w:t>caption gamma delta</w:t></w:r></w:p>"
MARKER = """<w:p>
  <w:pPr><w:spacing w:before="240"/></w:pPr>
  <w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t>MARKER HEADING</w:t></w:r>
</w:p>"""
SECTION = """<w:sectPr>
  <w:pgSz w:w="11907" w:h="16840"/>
  <w:pgMar w:top="1440" w:right="1134" w:bottom="1134" w:left="1418" w:header="709" w:footer="709"/>
</w:sectPr>"""


def document(side, width, empty_anchor):
    rows = "".join(
        f'<w:tr><w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>{content}</w:tc></w:tr>'
        for content in (PICTURE, CAPTION)
    )
    anchor = f'<w:p><w:r><w:t xml:space="preserve">{ANCHOR}</w:t></w:r></w:p>'
    paragraphs = "<w:p/>" + anchor + "<w:p/>" if empty_anchor else anchor + "<w:p/>"
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblpPr w:horzAnchor="text" w:vertAnchor="text" w:tblpXSpec="{side}" w:tblpY="1" w:leftFromText="141" w:rightFromText="141"/>
        <w:tblOverlap w:val="never"/>
        <w:tblW w:w="0" w:type="auto"/>
        <w:tblLayout w:type="fixed"/>
        <w:tblCellMar>
          <w:top w:w="28" w:type="dxa"/><w:left w:w="28" w:type="dxa"/>
          <w:bottom w:w="28" w:type="dxa"/><w:right w:w="28" w:type="dxa"/>
        </w:tblCellMar>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="{width}"/></w:tblGrid>
      {rows}
    </w:tbl>
    {paragraphs}{MARKER}{SECTION}
  </w:body>
</w:document>
"""


def generate():
    assert len(ANCHOR) == 350
    for name, side, width, empty_anchor in FIXTURES:
        parts = (
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", PACKAGE_RELS),
            ("word/_rels/document.xml.rels", DOCUMENT_RELS),
            ("word/styles.xml", STYLES),
            ("word/document.xml", document(side, width, empty_anchor)),
            ("word/media/pixel.png", PIXEL),
        )
        with ZipFile(ROOT / f"{name}.docx", "w", compression=ZIP_STORED) as archive:
            for path, content in parts:
                entry = ZipInfo(path, date_time=(1980, 1, 1, 0, 0, 0))
                entry.compress_type = ZIP_STORED
                entry.create_system = 0
                entry.external_attr = 0o600 << 16
                archive.writestr(entry, content.encode("utf-8") if isinstance(content, str) else content)


if __name__ == "__main__":
    generate()
