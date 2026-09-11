#!/usr/bin/env python3
"""Generate the deterministic redaction integrity DOCX fixture."""
from base64 import b64encode
from struct import pack
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo
import zlib


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "redaction-integrity.docx"
def make_png():
    def chunk(kind, data):
        return pack(">I", len(data)) + kind + data + pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", pack(">IIBBBBB", 1, 1, 8, 6, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(b"\x00\x00\x00\x00\x00")) + chunk(b"IEND", b"")


PNG = make_png()


def fixed_info(name):
    info = ZipInfo(name, (2000, 1, 1, 0, 0, 0))
    info.compress_type = ZIP_DEFLATED
    return info


def nested_gfxdata():
    import io

    payload = io.BytesIO()
    with ZipFile(payload, "w", ZIP_DEFLATED) as archive:
        archive.writestr(fixed_info("[Content_Types].xml"), "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>")
        archive.writestr(
            fixed_info("drs/e2oDoc.xml"),
            "<drs xmlns=\"urn:schemas-microsoft-com:office:drs\"><text>SECRET_GFX_TEXT</text></drs>",
        )
        archive.writestr(fixed_info("drs/media/image1.png"), PNG)
    return b64encode(payload.getvalue()).decode("ascii")


GFX = nested_gfxdata()
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"

parts = {
    "[Content_Types].xml": f'''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/customXml/item1.xml" ContentType="application/xml"/><Override PartName="/customXml/item2.xml" ContentType="application/xml"/><Override PartName="/customXml/item3.xml" ContentType="application/xml"/><Override PartName="/customXml/item4.xml" ContentType="application/xml"/>
<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/><Override PartName="/customXml/itemProps2.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/><Override PartName="/customXml/itemProps3.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/><Override PartName="/customXml/itemProps4.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>
</Types>''',
    "_rels/.rels": f'''<Relationships xmlns="{PR}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>''',
    "word/_rels/document.xml.rels": f'''<Relationships xmlns="{PR}"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdCustom1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/><Relationship Id="rIdCustom2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item2.xml"/><Relationship Id="rIdCustom3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item3.xml"/><Relationship Id="rIdCustom4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item4.xml"/><Relationship Id="rIdMedia" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>''',
    "word/styles.xml": f'''<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style><w:style w:type="paragraph" w:customStyle="1" w:styleId="AAAA1111"><w:name w:val="ABCD1234"/></w:style><w:style w:type="paragraph" w:customStyle="1" w:styleId="BBBB2222"><w:name w:val="WXYZ9876"/></w:style></w:styles>''',
    "word/document.xml": f'''<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body><w:p><w:pPr><w:pStyle w:val="AAAA1111"/></w:pPr><w:r><w:t>SECRET_VISIBLE_TEXT</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="BBBB2222"/></w:pPr><w:r><w:t>SECRET_SECOND_TEXT</w:t></w:r></w:p><w:p><w:r><w:t>SECRET_SHAPE_TEXT</w:t></w:r><w:r><w:pict><v:shape id="shape1" style="width:10pt;height:10pt" o:spid="_x0000_s1025" o:gfxdata="{GFX}"/></w:pict></w:r></w:p><w:sectPr/></w:body></w:document>''',
    "word/media/image1.png": PNG,
    "customXml/item1.xml": '<root xmlns="urn:synthetic:private"><record secret="SECRET_ATTR_VALUE">SECRET_PRIVATE_TEXT</record></root>',
    "customXml/item2.xml": '<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:tns="urn:synthetic:schema" xmlns:meta="urn:synthetic:metadata" meta:schemaLocation="SECRET_SCHEMA_LOCATION" targetNamespace="urn:synthetic:schema" elementFormDefault="qualified"><xs:annotation><xs:documentation>SECRET_ANNOTATION</xs:documentation></xs:annotation><xs:complexType name="RecordType"><xs:sequence><xs:element ref="tns:Code" minOccurs="1" maxOccurs="1"/><xs:element name="Nullable" type="xs:string" minOccurs="0" nillable="true"/></xs:sequence></xs:complexType><xs:element name="Record" type="tns:RecordType"/><xs:element name="Code" type="tns:CodeType"/><xs:element name="DefaultCode" type="xs:string" default="SECRET_DEFAULT"/><xs:element name="FixedCode" type="xs:string" fixed="SECRET_FIXED"/><xs:simpleType name="CodeType"><xs:restriction base="xs:string"><xs:maxLength value="12"/><xs:enumeration value="SECRET_ENUM"/><xs:pattern value="SECRET_[A-Z]+"/></xs:restriction></xs:simpleType></xs:schema>',
    "customXml/item3.xml": '<config xmlns="urn:synthetic:config" schemaLocation="SECRET_SCHEMA_LOCATION" default="SECRET_DEFAULT" fixed="SECRET_FIXED" enumeration="SECRET_ENUM" pattern="SECRET_PATTERN"/>',
    "customXml/item4.xml": '<data xmlns="urn:synthetic:data"><value>SECRET_OTHER_PRIVATE</value></data>',
}

for i in range(1, 5):
    parts[f"customXml/itemProps{i}.xml"] = f'<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{{00000000-0000-0000-0000-00000000000{i}}}"/>'

for i in range(1, 5):
    parts[f"customXml/_rels/item{i}.xml.rels"] = f'<Relationships xmlns="{PR}"><Relationship Id="rIdProps{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps{i}.xml"/></Relationships>'


def write_fixture():
    with ZipFile(OUT, "w", ZIP_DEFLATED) as archive:
        for name in sorted(parts):
            archive.writestr(fixed_info(name), parts[name])


if __name__ == "__main__":
    write_fixture()
