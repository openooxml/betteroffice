import io
import json
import zipfile
from pathlib import Path
from typing import Any, Callable

import pytest

ROOT = Path(__file__).resolve().parents[3]
FIXTURES = ROOT / "apps" / "demo" / "public"
LAYOUT_FIXTURES = ROOT / "crates" / "docx-layout" / "tests" / "fixtures"
FONTS = ROOT / "crates" / "docx-raster" / "tests" / "assets"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

PARTS = {
    "[Content_Types].xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        "</Types>"
    ),
    "_rels/.rels": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    ),
    "word/_rels/document.xml.rels": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rIdHeader" Type="{R}/header" Target="header1.xml"/>'
        "</Relationships>"
    ),
    "word/document.xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}" xmlns:w14="{W14}" xmlns:r="{R}"><w:body>'
        '<w:p w14:paraId="11111111"><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/>'
        '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr></w:pPr>'
        '<w:r><w:rPr><w:b/><w:sz w:val="48"/><w:color w:val="FF0000"/>'
        '<w:rFonts w:ascii="Calibri"/></w:rPr><w:t>Hello DOCX</w:t></w:r></w:p>'
        '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="1200"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'
        '<w:p w14:paraId="22222222"><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:tcPr><w:tcW w:w="1200" w:type="dxa"/></w:tcPr>'
        '<w:p w14:paraId="33333333"><w:r><w:t>Right</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
        '<w:p w14:paraId="44444444"><w:r><w:t xml:space="preserve">Plain </w:t></w:r>'
        "<w:r><w:rPr><w:i/></w:rPr><w:t>and italic</w:t></w:r></w:p>"
        '<w:p w14:paraId="55555555"><w:r><w:t>Second section</w:t></w:r></w:p>'
        '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        "</w:sectPr></w:body></w:document>"
    ),
    "word/header1.xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W}" xmlns:w14="{W14}">'
        '<w:p w14:paraId="66666666"><w:r><w:t>Native header</w:t></w:r></w:p></w:hdr>'
    ),
}


def _read(path: Path) -> bytes:
    if not path.exists():
        pytest.skip(f"fixture missing: {path}")
    return path.read_bytes()


def _package(parts: "dict[str, str]") -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, body in parts.items():
            archive.writestr(name, body)
    return buffer.getvalue()


@pytest.fixture(scope="session")
def minimal_bytes() -> bytes:
    """A document with paragraph IDs, a table, a header, and two sections."""
    return _package(PARTS)


@pytest.fixture(scope="session")
def nested_table_bytes() -> bytes:
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}"><w:body><w:tbl><w:tr><w:tc>'
        '<w:p><w:r><w:t>Outer cell</w:t></w:r></w:p>'
        '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Direct nested</w:t></w:r></w:p>'
        '</w:tc></w:tr></w:tbl>'
        '<w:sdt><w:sdtPr><w:alias w:val="Nested control"/></w:sdtPr><w:sdtContent>'
        '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>SDT nested</w:t></w:r></w:p>'
        '</w:tc></w:tr></w:tbl></w:sdtContent></w:sdt>'
        '</w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>'
    )
    return _package({**PARTS, "word/document.xml": document})


@pytest.fixture(scope="session")
def duplicate_id_package() -> "Callable[..., bytes]":
    """Builds a package as Word writes it: `1A2B3C4D` repeats in the body, a
    nested table cell and a content control, and the last paragraph's ID is
    `last_id` as authored, `0B000003` with a character reference by default."""

    def paragraph(para_id: str, text: str) -> str:
        return (
            f'<w:p w14:paraId="{para_id}" w14:textId="77777777">'
            f"<w:r><w:t>{text}</w:t></w:r></w:p>"
        )

    def cell(content: str) -> str:
        return f"<w:tbl><w:tr><w:tc>{content}</w:tc></w:tr></w:tbl>"

    def build(last_id: str = "0B00000&#x33;") -> bytes:
        document = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:document xmlns:w="{W}" xmlns:w14="{W14}" xmlns:r="{R}"><w:body>'
            f'{paragraph("1A2B3C4D", "first")}{paragraph("1A2B3C4D", "second")}'
            f'{cell(cell(paragraph("1A2B3C4D", "nested")) + paragraph("0B000002", ""))}'
            '<w:sdt><w:sdtPr><w:alias w:val="Clause"/></w:sdtPr><w:sdtContent>'
            f'{paragraph("1A2B3C4D", "control")}</w:sdtContent></w:sdt>'
            f'{paragraph(last_id, "third")}'
            '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr>'
            "</w:body></w:document>"
        )
        return _package({**PARTS, "word/document.xml": document})

    return build


@pytest.fixture(scope="session")
def duplicate_id_bytes(duplicate_id_package: "Callable[..., bytes]") -> bytes:
    return duplicate_id_package()


@pytest.fixture(scope="session")
def sample_path() -> Path:
    path = FIXTURES / "betteroffice-demo.docx"
    if not path.exists():
        pytest.skip(f"fixture missing: {path}")
    return path


@pytest.fixture(scope="session")
def sample_bytes(sample_path: Path) -> bytes:
    return sample_path.read_bytes()


@pytest.fixture(scope="session")
def layout_input() -> "dict[str, Any]":
    path = LAYOUT_FIXTURES / "single-page-multi-paragraph.input.json"
    return json.loads(_read(path))


@pytest.fixture(scope="session")
def font_bytes() -> bytes:
    return _read(FONTS / "Carlito-Regular.ttf")
