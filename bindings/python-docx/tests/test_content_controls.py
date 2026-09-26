from pathlib import Path

import pytest

from betteroffice_docx import Document, ExportError

TEMPLATE = (
    Path(__file__).resolve().parents[3]
    / "packages"
    / "docx"
    / "src"
    / "yrs"
    / "__fixtures__"
    / "content-controls"
    / "template.docx"
)
KEYS = [
    "controlId",
    "ooxmlId",
    "controlType",
    "tag",
    "alias",
    "lock",
    "showingPlaceholder",
    "dataBound",
    "placement",
    "anchor",
    "parentControlId",
    "value",
    "multiLine",
    "effectiveLock",
]


@pytest.fixture(scope="module")
def template_bytes() -> bytes:
    return TEMPLATE.read_bytes()


def test_lists_every_control_with_the_engine_dto(template_bytes: bytes) -> None:
    document = Document.open(template_bytes)
    snapshot = document.list_content_controls()

    assert snapshot["schemaVersion"] == 1
    assert snapshot["anchorScope"] == "snapshot"
    assert snapshot["complete"] is True
    assert [control["tag"] for control in snapshot["controls"]] == [
        "customer.name",
        "account.reference",
        "account.reference",
        "terms.standard",
        "customer.address",
        "customer.email",
        "document.title",
    ]
    name = snapshot["controls"][0]
    assert list(name) == KEYS
    assert name["value"] == {"kind": "text", "text": "Click to enter a name."}
    assert name["showingPlaceholder"] is True
    assert snapshot["controls"][4]["placement"] == "block"
    assert snapshot["controls"][4]["value"]["text"] == "1 Old Road\nOldtown"
    assert snapshot["controls"][3]["effectiveLock"] == {"content": True, "control": False, "known": True}
    assert snapshot["controls"][5]["dataBound"] is True


def test_find_returns_every_exact_match(template_bytes: bytes) -> None:
    document = Document.open(template_bytes)
    duplicates = document.find_content_controls({"kind": "tag", "tag": "account.reference"})
    ids = [control["controlId"] for control in duplicates["controls"]]

    assert len(ids) == 2 and len(set(ids)) == 2
    assert document.find_content_controls({"kind": "tag", "tag": "ACCOUNT.reference"})["controls"] == []
    assert len(document.find_content_controls({"kind": "alias", "alias": "Address"})["controls"]) == 1
    by_ooxml_id = document.find_content_controls({"kind": "ooxmlId", "ooxmlId": "105"})
    assert [control["tag"] for control in by_ooxml_id["controls"]] == ["customer.address"]
    headers = document.list_content_controls(stories=["headers"])
    assert [control["tag"] for control in headers["controls"]] == ["document.title"]


def test_reads_leave_the_document_unchanged(template_bytes: bytes) -> None:
    document = Document.open(template_bytes)
    saved = document.save()
    document.list_content_controls()
    document.find_content_controls({"kind": "tag", "tag": "customer.name"})

    assert document.save() == saved


def test_refusals_raise_with_the_failure(template_bytes: bytes) -> None:
    document = Document.open(template_bytes)
    with pytest.raises(ExportError, match="limit-exceeded") as limited:
        document.list_content_controls(max_controls=2)
    assert limited.value.failure["code"] == "limit-exceeded"
    with pytest.raises(ExportError, match="invalid-options"):
        document.list_content_controls(max_controls=0)
    with pytest.raises(ValueError, match="query"):
        document.find_content_controls({"kind": "name", "name": "x"})
    with pytest.raises(ValueError, match="stories"):
        document.list_content_controls(stories=["pages"])
