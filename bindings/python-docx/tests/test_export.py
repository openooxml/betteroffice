import json
from pathlib import Path

import pytest

from betteroffice_docx import Document, ExportError, render_docx_markdown

EXPORT_FIXTURES = (
    Path(__file__).resolve().parents[3] / "crates" / "docx-edit" / "tests" / "fixtures" / "structured-export"
)
ALL_STORIES = ["body", "headers", "footers", "footnotes", "endnotes", "comments"]


@pytest.fixture(scope="module")
def principal() -> Document:
    return Document.open((EXPORT_FIXTURES / "principal.docx").read_bytes())


@pytest.mark.parametrize("view", ["accepted", "original", "markup"])
def test_export_matches_the_engine_golden_files(principal: Document, view: str) -> None:
    content = principal.export_structured(revision_view=view, stories=ALL_STORIES)
    expected = json.loads((EXPORT_FIXTURES / f"principal.{view}.json").read_text())

    assert content == expected
    assert content["schemaVersion"] == 1
    assert content["anchorScope"] == "snapshot"

    markdown = principal.export_markdown(revision_view=view, stories=ALL_STORIES)
    assert markdown["markdown"] == (EXPORT_FIXTURES / f"principal.{view}.md").read_text()
    assert markdown == render_docx_markdown(content)
    assert len(markdown["anchors"]) == markdown["markdown"].count("<!-- docx-export:")


def test_export_defaults_to_the_body_and_lists_what_it_leaves_out(principal: Document) -> None:
    content = principal.export_structured(revision_view="accepted")

    assert [story["story"] for story in content["stories"]] == ["body"]
    assert content["includedStories"] == ["body"]
    omitted = [d for d in content["diagnostics"] if d["code"] == "stories-omitted"]
    assert len(omitted) == 5


def test_export_reads_edits_made_through_the_document(minimal_bytes: bytes) -> None:
    document = Document.open(minimal_bytes)
    document.replace_text("11111111", "Edited in Python")

    content = document.export_structured(revision_view="accepted", include_formatting=False)
    first = content["stories"][0]["blocks"][0]

    assert first["anchor"] == {"kind": "paragraph", "story": "body", "paraId": "11111111"}
    assert [inline["text"] for inline in first["paragraph"]["inlines"]] == ["Edited in Python"]
    assert all(inline["marks"] is None for inline in first["paragraph"]["inlines"])


def test_content_without_a_location_is_anchored_unlocated(principal: Document) -> None:
    content = principal.export_structured(revision_view="accepted")
    unlocated = {"kind": "unlocated", "story": "body", "reason": "duplicate-paragraph-id"}
    content["stories"][0]["blocks"][0]["anchor"] = unlocated

    markdown = render_docx_markdown(content)

    assert markdown["anchors"][0]["anchor"] == unlocated


def test_unusable_options_raise(principal: Document) -> None:
    with pytest.raises(ExportError, match="invalid-options") as refused:
        principal.export_structured(revision_view="accepted", max_bytes=16)
    assert refused.value.failure["code"] == "invalid-options"
    assert refused.value.failure["target"] is None
    with pytest.raises(ExportError, match="limit-exceeded") as limited:
        principal.export_structured(revision_view="accepted", max_blocks=2_000_000)
    assert limited.value.failure["code"] == "limit-exceeded"
    with pytest.raises(ValueError, match="revision_view"):
        principal.export_structured(revision_view="final")
    with pytest.raises(ValueError, match="stories"):
        principal.export_structured(revision_view="accepted", stories=["pages"])


def test_truncated_exports_stay_within_their_limits(principal: Document) -> None:
    content = principal.export_structured(
        revision_view="markup", stories=ALL_STORIES, max_bytes=4_096
    )

    assert content["truncated"] is True
    assert content["diagnostics"][-1]["code"] == "truncated"
    assert len(json.dumps(content, separators=(",", ":"), ensure_ascii=False).encode()) <= 4_096


def test_rendering_rejects_other_schema_versions(principal: Document) -> None:
    content = principal.export_structured(revision_view="accepted")
    content["schemaVersion"] = 2

    with pytest.raises(ValueError, match="schema version"):
        render_docx_markdown(content)
