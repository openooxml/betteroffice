from pathlib import Path
from xml.sax.saxutils import escape
import zipfile


W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
STYLES = f'''<w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/><w:widowControl w:val="0"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'''
BREAK = '<w:r><w:br w:type="page"/></w:r>'


def paragraph(text, properties="", prefix=""):
    return f'<w:p><w:pPr>{properties}</w:pPr>{prefix}<w:r><w:t>{escape(text)}</w:t></w:r></w:p>'


def document(name, body, compat="", section_extra=""):
    section = f'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>{section_extra}</w:sectPr>'
    parts = {
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
        "_rels/.rels": f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>',
        "word/_rels/document.xml.rels": f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="{REL}/styles" Target="styles.xml"/><Relationship Id="settings" Type="{REL}/settings" Target="settings.xml"/></Relationships>',
        "word/document.xml": f'<w:document {W}><w:body>{body}{section}</w:body></w:document>',
        "word/styles.xml": STYLES,
        "word/settings.xml": f'<w:settings {W}><w:compat>{compat}<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat></w:settings>',
    }
    destination = Path(__file__).parent / f"{name}.docx"
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as archive:
        for key, value in parts.items():
            info = zipfile.ZipInfo(key, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, value)


def main():
    auto = (
        paragraph("ANCHOR")
        + paragraph("AUTO BEFORE", '<w:spacing w:before="100" w:beforeAutospacing="1" w:after="0"/>')
        + paragraph("AUTO AFTER", '<w:spacing w:before="0" w:after="200" w:afterAutospacing="1"/>')
        + paragraph("AFTER TARGET")
        + paragraph("AUTO BOTH", '<w:spacing w:before="100" w:after="200" w:beforeAutospacing="1" w:afterAutospacing="1"/>')
        + paragraph("BOTH TARGET")
    )
    breaks = (
        paragraph("FILLER")
        + paragraph("HARD TARGET", '<w:spacing w:before="480"/>', BREAK)
        + paragraph("PROPERTY TARGET", '<w:pageBreakBefore/><w:spacing w:before="480"/>')
        + f"<w:p>{BREAK}</w:p>"
        + paragraph("STANDALONE TARGET", '<w:spacing w:before="480"/>')
    )
    contextual = '<w:contextualSpacing/><w:spacing w:before="240" w:after="360"/>'
    cell = (
        paragraph("CELL 1", '<w:contextualSpacing/><w:spacing w:before="0" w:after="360"/>')
        + paragraph("CELL 2", contextual)
        + paragraph("CELL 3", '<w:contextualSpacing/><w:spacing w:before="240" w:after="0"/>')
    )
    table = (
        paragraph("BODY 1", contextual)
        + paragraph("BODY 2", '<w:contextualSpacing/><w:spacing w:before="240" w:after="0"/>')
        + '<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>'
        + cell
        + '</w:tc></w:tr></w:tbl>'
        + paragraph("TABLE END")
    )
    for name, body, flag in [
        ("auto", auto, "doNotUseHTMLParagraphAutoSpacing"),
        ("break", breaks, "suppressSpBfAfterPgBrk"),
        ("table", table, "allowSpaceOfSameStyleInTable"),
    ]:
        document(f"{name}-default", body)
        document(f"{name}-enabled", body, f"<w:{flag}/>")
    probes = (
        paragraph("ANCHOR")
        + paragraph("BEFORE ONLY", '<w:spacing w:beforeAutospacing="1"/>')
        + paragraph("AFTER ONLY", '<w:spacing w:afterAutospacing="1"/>')
        + paragraph("AFTER TARGET")
        + paragraph("BOTH", '<w:spacing w:beforeAutospacing="1" w:afterAutospacing="1"/>')
        + paragraph("BOTH TARGET")
    )
    document("auto-without-fallback", probes, "<w:doNotUseHTMLParagraphAutoSpacing/>")


if __name__ == "__main__":
    main()
