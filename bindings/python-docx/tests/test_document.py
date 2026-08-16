from pathlib import Path

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
