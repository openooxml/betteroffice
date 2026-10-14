import io
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from betteroffice_docx import Document
from conftest import PARTS, R, W, _package

FOREIGN = "urn:betteroffice-python-test"


def _paragraph(text: str) -> str:
    return f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"


def _raw_block(label: str) -> str:
    return (
        f'<bofx:opaque xmlns:bofx="{FOREIGN}" bofx:label="{label}" '
        'bofx:value="A&amp;B&lt;C">'
        f'{_paragraph("Hidden paragraph")}'
        f'<w:tbl><w:tr><w:tc>{_paragraph("Hidden cell")}</w:tc></w:tr></w:tbl>'
        "</bofx:opaque>"
    )


@pytest.fixture
def raw_xml_bytes() -> bytes:
    document = (
        f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>'
        f'{_raw_block("body")}{_paragraph("Body")}'
        "<w:sdt><w:sdtPr/><w:sdtContent>"
        f'{_raw_block("sdt")}{_paragraph("Control")}'
        "</w:sdtContent></w:sdt><w:tbl><w:tr><w:tc>"
        f'{_raw_block("cell")}{_paragraph("Cell")}'
        f'<w:tbl><w:tr><w:tc>{_paragraph("Nested")}</w:tc></w:tr></w:tbl>'
        "</w:tc></w:tr></w:tbl><w:sectPr>"
        '<w:headerReference w:type="default" r:id="rIdHeader"/>'
        "</w:sectPr></w:body></w:document>"
    )
    header = (
        f'<w:hdr xmlns:w="{W}">{_raw_block("header")}<w:p>'
        f'<bofx:marker xmlns:bofx="{FOREIGN}" bofx:value="A&amp;B&lt;C"/>'
        "<w:r><w:t>Header</w:t></w:r></w:p></w:hdr>"
    )
    return _package({**PARTS, "word/document.xml": document, "word/header1.xml": header})


def test_raw_blocks_are_opaque_to_typed_enumeration(raw_xml_bytes: bytes) -> None:
    document = Document.open(raw_xml_bytes)
    texts = ["Body", "Control", "Cell", "Nested"]

    assert [paragraph.text for paragraph in document.paragraphs()] == texts
    assert [paragraph.text for paragraph in document.sections()[0].paragraphs] == texts
    outer, nested = document.tables()
    (cell,) = outer.rows[0].cells
    assert [paragraph.text for paragraph in cell.paragraphs] == ["Cell", "Nested"]
    assert [table.text for table in cell.tables] == [nested.text] == ["Nested"]
    (header,) = document.headers()
    assert [paragraph.text for paragraph in header.paragraphs] == ["Header"]
    assert [run.text for run in header.paragraphs[0].runs] == ["Header"]
    assert (document.structure().body_paragraphs, document.structure().body_tables) == (4, 2)


def test_raw_block_and_inline_xml_survive_save(
    raw_xml_bytes: bytes, tmp_path: Path
) -> None:
    (tmp_path / "raw-input.docx").write_bytes(raw_xml_bytes)
    saved = tmp_path / "raw-saved.docx"
    Document.open(raw_xml_bytes).save_path(saved)

    with zipfile.ZipFile(io.BytesIO(raw_xml_bytes)) as original, zipfile.ZipFile(saved) as archive:
        parsed = {
            name: ET.fromstring(archive.read(name))
            for name in archive.namelist()
            if name.endswith((".xml", ".rels"))
        }
        for part, count in (("word/document.xml", 3), ("word/header1.xml", 1)):
            before = ET.fromstring(original.read(part)).findall(f".//{{{FOREIGN}}}opaque")
            after = parsed[part].findall(f".//{{{FOREIGN}}}opaque")
            assert len(after) == count
            assert [ET.tostring(node) for node in after] == [ET.tostring(node) for node in before]
        marker = parsed["word/header1.xml"].find(f".//{{{FOREIGN}}}marker")
        assert marker is not None
        assert marker.attrib[f"{{{FOREIGN}}}value"] == "A&B<C"

    assert Document.open_path(saved).save() == saved.read_bytes()
