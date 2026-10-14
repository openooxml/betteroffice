import html
import io
import re
import zipfile
from pathlib import Path
from typing import Callable

import pytest

from betteroffice_docx import (
    Document,
    DocxError,
    ParseError,
    UnsupportedEditError,
    __version__,
)


def test_version_is_exposed() -> None:
    assert __version__.count(".") == 2


def test_reads_the_body_structure(minimal_bytes: bytes) -> None:
    document = Document.open(minimal_bytes)
    structure = document.structure()

    assert (structure.body_paragraphs, structure.body_tables) == (5, 1)
    assert (structure.sections, structure.headers) == (2, 1)
    assert len(document) == 5
    assert document.paragraph_ids == [
        "11111111",
        "22222222",
        "33333333",
        "44444444",
        "55555555",
    ]
    assert document.text.splitlines() == [
        "Hello DOCX",
        "Cell text",
        "Right",
        "Plain and italic",
        "Second section",
    ]


def test_paragraph_carries_style_alignment_and_run_formatting(
    minimal_bytes: bytes,
) -> None:
    heading = Document.open(minimal_bytes).paragraph("11111111")

    assert (heading.style, heading.alignment) == ("Heading1", "center")
    (run,) = heading.runs
    assert run.text == "Hello DOCX"
    assert run.bold is True
    assert run.italic is None
    assert run.font_size == 24.0
    assert run.color == "FF0000"
    assert run.font_family == "Calibri"


def test_run_texts_concatenate_to_the_paragraph_text(minimal_bytes: bytes) -> None:
    for paragraph in Document.open(minimal_bytes):
        assert "".join(run.text for run in paragraph.runs) == paragraph.text


def test_table_cells_are_reachable_through_the_table_and_the_body(
    minimal_bytes: bytes,
) -> None:
    document = Document.open(minimal_bytes)
    (table,) = document.tables()

    assert len(table) == 1
    assert table.column_widths == [2400.0, 1200.0]
    assert [cell.text for cell in table.rows[0].cells] == ["Cell text", "Right"]
    assert len(table.rows[0]) == 2
    assert table.rows[0].cells[0].tables == []
    assert document.paragraph("22222222").text == "Cell text"


def test_nested_tables_are_reachable_from_their_cell(nested_table_bytes: bytes) -> None:
    document = Document.open(nested_table_bytes)
    outer, direct, controlled = document.tables()
    (cell,) = outer.rows[0].cells

    assert [table.text for table in cell.tables] == ["Direct nested", "SDT nested"]
    assert [table.text for table in document.tables()[1:]] == [direct.text, controlled.text]
    assert len(document.tables()) == len(cell.tables) + 1


def test_sections_and_headers(minimal_bytes: bytes) -> None:
    document = Document.open(minimal_bytes)
    first, second = document.sections()

    assert first.text == "Hello DOCX"
    assert (second.page_width, second.page_height) == (12240.0, 15840.0)
    assert second.margin_left == 1440.0
    assert [paragraph.text for paragraph in second.paragraphs][-1] == "Second section"

    (header,) = document.headers()
    assert (header.rel_id, header.kind, header.text) == (
        "rIdHeader",
        "default",
        "Native header",
    )
    assert document.footers() == []


def test_paragraph_lookup_by_id_and_index(minimal_bytes: bytes) -> None:
    document = Document.open(minimal_bytes)

    assert document[0].text == "Hello DOCX"
    assert document["55555555"].text == "Second section"
    assert document.paragraph(3).id == "44444444"

    with pytest.raises(KeyError):
        document.paragraph("deadbeef")
    with pytest.raises(IndexError):
        document.paragraph(99)
    with pytest.raises(TypeError):
        document.paragraph(True)


def test_replace_text_round_trips_through_save(minimal_bytes: bytes) -> None:
    document = Document.open(minimal_bytes)
    before = document.structure()

    edit = document.replace_text("11111111", "Edited from Python")
    assert (edit.para_id, edit.story) == ("11111111", "body")
    # The receipt carries the resulting range, so it spans the new text.
    assert (edit.start, edit.end) == (0, len("Edited from Python"))

    reopened = Document.open(document.save())
    after = reopened.structure()

    assert reopened.paragraph("11111111").text == "Edited from Python"
    assert (after.body_paragraphs, after.body_tables) == (
        before.body_paragraphs,
        before.body_tables,
    )
    assert (after.sections, after.headers) == (before.sections, before.headers)
    assert reopened.paragraph("22222222").text == "Cell text"
    assert reopened.headers()[0].text == "Native header"


def test_a_rewritten_paragraph_keeps_its_style_and_run_formatting(
    minimal_bytes: bytes,
) -> None:
    document = Document.open(minimal_bytes)
    document.replace_text("11111111", "Still a heading")

    heading = Document.open(document.save()).paragraph("11111111")
    (run,) = heading.runs
    assert (heading.style, heading.alignment) == ("Heading1", "center")
    assert (run.text, run.bold, run.font_size, run.color) == (
        "Still a heading",
        True,
        24.0,
        "FF0000",
    )


def test_replace_text_refuses_what_it_cannot_rebuild(minimal_bytes: bytes) -> None:
    document = Document.open(minimal_bytes)

    with pytest.raises(KeyError):
        document.replace_text("deadbeef", "nope")
    with pytest.raises(UnsupportedEditError) as raised:
        document.replace_text("44444444", "nope")
    assert issubclass(raised.type, DocxError)

    # A refused edit leaves the document intact.
    assert Document.open(document.save()).paragraph("44444444").text == (
        "Plain and italic"
    )


REPEATED = "1A2B3C4D"


def _written_ids(package: bytes) -> "dict[str, list[str]]":
    """The unescaped `w14:paraId` of every `w:p` in each XML part."""
    with zipfile.ZipFile(io.BytesIO(package)) as archive:
        return {
            name: [
                html.unescape(value)
                for value in re.findall(
                    r'<w:p\b[^>]*?w14:paraId="([^"]*)"', archive.read(name).decode()
                )
            ]
            for name in archive.namelist()
            if name.endswith(".xml")
        }


def _parts(package: bytes) -> "dict[str, bytes]":
    with zipfile.ZipFile(io.BytesIO(package)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def test_a_repeated_paragraph_id_addresses_its_own_paragraph(
    duplicate_id_bytes: bytes,
) -> None:
    document = Document.open(duplicate_id_bytes)
    ids = document.paragraph_ids

    assert (ids[0], ids[3], ids[5]) == (REPEATED, "0B000002", "0B000003")
    assert len(set(ids)) == 6
    assert [document[id].text for id in ids] == [
        "first",
        "second",
        "nested",
        "",
        "control",
        "third",
    ]

    assert document.replace_text(ids[1], "replacement").para_id == ids[1]

    reopened = Document.open(document.save())
    assert [(paragraph.id, paragraph.text) for paragraph in reopened.paragraphs()] == [
        (REPEATED, "first"),
        (ids[1], "replacement"),
        (ids[2], "nested"),
        ("0B000002", ""),
        (ids[4], "control"),
        ("0B000003", "third"),
    ]


def test_an_unedited_save_keeps_every_other_part_and_authored_paragraph_id(
    duplicate_id_bytes: bytes,
) -> None:
    document = Document.open(duplicate_id_bytes)
    ids = document.paragraph_ids
    assert REPEATED not in ids[1:]
    for para_id in ids:
        assert document.paragraph(para_id).id == para_id

    saved = document.save()
    before, after = _parts(duplicate_id_bytes), _parts(saved)
    assert list(after) == list(before)
    serialized = {"word/document.xml", "word/header1.xml"}
    for name in before.keys() - serialized:
        assert after[name] == before[name], name
    written, authored = _written_ids(saved), _written_ids(duplicate_id_bytes)
    for name in serialized:
        assert written[name] == authored[name], name
    assert written["word/document.xml"] == [
        REPEATED,
        REPEATED,
        REPEATED,
        "0B000002",
        REPEATED,
        "0B000003",
    ]


def test_repeated_ids_stay_addressable_across_save_and_reopen_cycles(
    duplicate_id_bytes: bytes,
) -> None:
    ids = Document.open(duplicate_id_bytes).paragraph_ids
    first = Document.open(duplicate_id_bytes).save()
    assert Document.open(first).paragraph_ids == ids
    assert Document.open(first).save() == first

    cycled = duplicate_id_bytes
    for index, text in [(2, "nested edit"), (4, "control edit"), (1, "second edit")]:
        document = Document.open(cycled)
        assert document.paragraph_ids == ids
        document.replace_text(ids[index], text)
        cycled = document.save()

    reopened = Document.open(cycled)
    assert [(paragraph.id, paragraph.text) for paragraph in reopened.paragraphs()] == [
        (REPEATED, "first"),
        (ids[1], "second edit"),
        (ids[2], "nested edit"),
        ("0B000002", ""),
        (ids[4], "control edit"),
        ("0B000003", "third"),
    ]
    assert reopened.save() == cycled
    assert _written_ids(cycled)["word/document.xml"] == ids


def test_a_fresh_id_avoids_an_id_written_with_a_character_reference(
    duplicate_id_package: "Callable[..., bytes]",
) -> None:
    taken = Document.open(duplicate_id_package()).paragraph_ids[1]
    escaped = f"{taken[:7]}&#x{ord(taken[7]):X};"
    package = duplicate_id_package(escaped)

    ids = Document.open(package).paragraph_ids
    assert ids[5] == taken
    assert taken not in ids[1:5]
    assert len(set(ids)) == 6
    assert _written_ids(Document.open(package).save())["word/document.xml"] == [
        REPEATED,
        REPEATED,
        REPEATED,
        "0B000002",
        REPEATED,
        taken,
    ]


def test_save_is_deterministic_and_save_path_writes_a_readable_file(
    minimal_bytes: bytes, tmp_path: Path
) -> None:
    document = Document.open(minimal_bytes)
    document.replace_text("55555555", "Written to disk")

    assert document.save() == document.save()

    target = tmp_path / "out.docx"
    document.save_path(target)
    assert Document.open_path(target).paragraph("55555555").text == "Written to disk"


def test_the_timestamp_is_the_only_clock(minimal_bytes: bytes) -> None:
    document = Document.open(minimal_bytes)
    assert document.timestamp == "1970-01-01T00:00:00.000Z"

    document.author = "ana"
    document.origin = "agent"
    document.timestamp = "2026-01-02T03:04:05.000Z"
    assert (document.author, document.origin) == ("ana", "agent")

    with pytest.raises(ValueError):
        document.origin = "nobody"


def test_parse_limits_reject_a_document_over_budget(minimal_bytes: bytes) -> None:
    with pytest.raises(ParseError):
        Document.open(minimal_bytes, limits={"max_paragraphs": 2})
    with pytest.raises(ValueError):
        Document.open(minimal_bytes, limits={"max_bananas": 2})

    assert len(Document.open(minimal_bytes, limits={"max_paragraphs": 500})) == 5


def test_unreadable_input_raises_parse_error() -> None:
    with pytest.raises(ParseError):
        Document.open(b"not a docx")


def test_open_accepts_every_bytes_like(minimal_bytes: bytes) -> None:
    assert len(Document.open(bytearray(minimal_bytes))) == 5
    assert len(Document.open(memoryview(minimal_bytes))) == 5
    with pytest.raises(TypeError):
        Document.open("not bytes")  # type: ignore[arg-type]


def test_document_is_usable_from_another_thread(minimal_bytes: bytes) -> None:
    import threading

    document = Document.open(minimal_bytes)
    results: list[str] = []
    thread = threading.Thread(target=lambda: results.append(document.text))
    thread.start()
    thread.join()

    assert results == [document.text]


def test_document_can_be_dropped_on_another_thread(
    minimal_bytes: bytes, monkeypatch: pytest.MonkeyPatch
) -> None:
    import sys
    import threading

    unraisable: list[object] = []
    monkeypatch.setattr(sys, "unraisablehook", unraisable.append)
    holder = [Document.open(minimal_bytes)]
    thread = threading.Thread(target=holder.clear)
    thread.start()
    thread.join()

    assert holder == []
    assert unraisable == []


def test_reads_the_demo_document(sample_path: Path) -> None:
    document = Document.open_path(sample_path)
    structure = document.structure()

    assert structure.body_paragraphs > 10
    assert structure.body_tables == 1
    assert document.paragraphs()[0].text == "Welcome to BetterOffice"
    assert document.paragraphs()[0].style == "Title"
    assert document.tables()[0].rows[0].cells[0].text == "Stage"
    # Word stamps no `w14:paraId` here, so nothing is addressable by ID.
    assert set(document.paragraph_ids) == {None}
    assert document.warnings == []
