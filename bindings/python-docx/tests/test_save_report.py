import io
import zipfile

import pytest

from betteroffice_docx import Document

from conftest import PARTS, R, W, W14, _package

SMART_TAG = (
    '<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place">'
    "<w:r><w:t>Ancient</w:t></w:r></w:smartTag>"
)
SMART_TAG_PARAGRAPH = (
    '<w:p w14:paraId="0000000B" w14:textId="77777777">'
    f'<w:r><w:t xml:space="preserve">Near </w:t></w:r>{SMART_TAG}</w:p>'
)
EMPTY_MARKED = (
    '<w:p w14:paraId="0000000E" w14:textId="77777777">'
    "<w:pPr><w:rPr><w:b/></w:rPr></w:pPr></w:p>"
)


def _paragraph(para_id: str, text: str) -> str:
    return (
        f'<w:p w14:paraId="{para_id}" w14:textId="77777777">'
        f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    )


@pytest.fixture(scope="module")
def report_bytes() -> bytes:
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}" xmlns:w14="{W14}" xmlns:r="{R}"><w:body>'
        f'{_paragraph("0000000A", "Alpha.")}{SMART_TAG_PARAGRAPH}{EMPTY_MARKED}'
        f'{_paragraph("0000000D", "Delta.")}'
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>'
    )
    core = (
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
        ' xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:modified>'
        "</cp:coreProperties>"
    )
    return _package({**PARTS, "word/document.xml": document, "docProps/core.xml": core})


def _main_part(package: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(package)) as archive:
        return archive.read("word/document.xml").decode()


def _parts(package: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(package)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _report(document: Document, **options: object):
    report = document.save_report(**options)
    assert report.bytes == document.save(**options)
    return report


def test_an_unedited_save_reports_nothing_kept_or_dropped(report_bytes: bytes) -> None:
    report = _report(Document.open(report_bytes))
    assert report.kept_unchanged is None
    assert report.structure_dropped is None
    assert report.whole_model_write is False
    assert report.bytes == report_bytes

    stamped = Document.open(report_bytes).save(
        update_modified_date=True, now="2026-01-02T03:04:05.000Z"
    )
    before, after = _parts(report_bytes), _parts(stamped)
    assert before.keys() == after.keys()
    assert [name for name in before if before[name] != after[name]] == ["docProps/core.xml"]
    assert b"2026-01-02T03:04:05.000Z" in after["docProps/core.xml"]


def test_one_edited_paragraph_reports_nothing_and_keeps_the_smart_tag(
    report_bytes: bytes,
) -> None:
    document = Document.open(report_bytes)
    document.replace_text("0000000D", "Delta edited.")
    report = _report(document)
    assert report.kept_unchanged is None
    assert report.structure_dropped is None
    assert report.whole_model_write is False
    saved = _main_part(report.bytes)
    assert "Delta edited." in saved
    assert SMART_TAG_PARAGRAPH in saved


def test_a_model_fallback_writes_the_edit_with_a_save_plan(report_bytes: bytes) -> None:
    document = Document.open(report_bytes)
    document.replace_text("0000000E", "X")
    report = _report(document)
    assert report.kept_unchanged is None
    assert report.whole_model_write is False
    assert report.structure_dropped is not None
    assert [(part.part, part.dropped) for part in report.structure_dropped] == [
        ("word/document.xml", ["w:smartTag"])
    ]
    assert any(
        warning.startswith("DOCX_PART_REWRITTEN word/document.xml ")
        for warning in report.warnings
    )
    assert ">X<" in _main_part(report.bytes)
    assert Document.open(report.bytes).paragraph("0000000E").text == "X"
    before, after = _parts(report_bytes), _parts(report.bytes)
    assert before.keys() == after.keys()
    assert [name for name in before if before[name] != after[name]] == ["word/document.xml"]


STRUCTURE_DROPPED_BODY = (
    '<w:p w14:paraId="0000000A" w14:textId="77777777"><w:r><w:t xml:space="preserve">Near </w:t></w:r>'
    f"{SMART_TAG}</w:p>"
    '<w:p w14:paraId="0000000B" w14:textId="77777777"><w:customXml w:uri="urn:pr1" w:element="tagged">'
    "<w:r><w:t>Tagged</w:t></w:r></w:customXml></w:p>"
    '<w:p w14:paraId="0000000C" w14:textId="77777777"><w:bookmarkStart w:id="1" w:name="pr1mark"/>'
    '<w:r><w:t>Marked</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>'
    '<w:p w14:paraId="0000000D" w14:textId="77777777"><w:pPr><mc:AlternateContent>'
    '<mc:Choice Requires="w14"><w:jc w:val="center"/></mc:Choice>'
    '<mc:Fallback><w:jc w:val="left"/></mc:Fallback></mc:AlternateContent></w:pPr>'
    "<w:r><w:t>aaaa</w:t></w:r></w:p>"
    '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
)


@pytest.fixture(scope="module")
def structure_dropped_bytes() -> bytes:
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
        f' xmlns:r="{R}" xmlns:w="{W}" xmlns:w14="{W14}" mc:Ignorable="w14"><w:body>'
        f"{STRUCTURE_DROPPED_BODY}"
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    return _package(
        {
            "[Content_Types].xml": content_types,
            "_rels/.rels": PARTS["_rels/.rels"],
            "word/document.xml": document,
        }
    )


def test_a_whole_write_in_the_save_plan_reports_the_structure_it_drops(
    structure_dropped_bytes: bytes,
) -> None:
    document = Document.open(structure_dropped_bytes)
    document.replace_text("0000000D", "aaa")
    report = _report(document)
    assert report.kept_unchanged is None
    assert report.structure_dropped is not None
    assert [(s.part, s.dropped) for s in report.structure_dropped] == [
        ("word/document.xml", ["w:smartTag", "w:customXml"])
    ]
    assert [
        (p.ordinal, p.para_id, p.text, p.dropped)
        for p in report.structure_dropped[0].paragraphs
    ] == [
        (0, "0000000A", "Near Ancient", ["w:smartTag"]),
        (1, "0000000B", "Tagged", ["w:customXml"]),
    ]


@pytest.mark.parametrize("edit", [None, "0000000D", "0000000E"])
def test_the_report_bytes_are_the_save_bytes_with_options(
    report_bytes: bytes, edit: "str | None"
) -> None:
    document = Document.open(report_bytes)
    if edit is not None:
        document.replace_text(edit, "Edited")
    options = {
        "now": "2026-01-02T03:04:05.000Z",
        "update_modified_date": True,
        "modified_by": "Reviewer",
    }
    assert document.save_report(**options).bytes == document.save(**options)
