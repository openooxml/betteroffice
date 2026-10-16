import pytest
import betteroffice_pptx as bo


def _paragraphs(shapes):
    for shape in shapes:
        for story in shape["stories"]:
            yield from story["paragraphs"]
        for row in (shape["table"] or {}).get("rows", []):
            for cell in row["cells"]:
                if cell["story"]:
                    yield from cell["story"]["paragraphs"]
        yield from _paragraphs(shape["children"])


def test_session_export_is_versioned_and_changes_nothing(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    version = deck.version()
    read = deck.export_structured(include_notes=True)
    assert read["ok"] and read["version"] == version
    content = read["content"]
    assert content["schemaVersion"] == 1
    assert content["anchorScope"] == "session"
    assert content["readingOrder"] == "shapeTree"
    assert content["included"]["notes"] is True
    assert [slide["index"] for slide in content["slides"]] == [0, 1, 2]

    stories = deck.read_content()["stories"]
    ranges = {
        paragraph["paragraphId"]: (paragraph["start"], paragraph["end"])
        for story in stories
        for paragraph in story["paragraphs"]
    }
    exported = [
        paragraph
        for slide in content["slides"]
        for paragraph in _paragraphs(slide["shapes"])
    ]
    assert exported
    for paragraph in exported:
        anchor = paragraph["anchor"]
        assert anchor["kind"] == "range"
        assert (anchor["start"], anchor["end"]) == ranges[paragraph["paragraphId"]]

    markdown = deck.export_markdown()
    assert markdown["ok"] and markdown["version"] == version
    assert markdown["content"]["markdown"].count("<!-- pptx-export:") == len(
        markdown["content"]["anchors"]
    )

    refused = deck.export_structured(max_bytes=8)
    assert refused == {
        "ok": False,
        "version": version,
        "failure": {
            "code": "invalid-options",
            "target": None,
            "message": "maxBytes must be at least 1024",
        },
    }
    assert deck.version() == version
    assert not deck.is_edited
    assert not deck.can_undo


def test_range_anchors_are_batch_targets(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    read = deck.export_structured()
    anchor = next(
        run["anchor"]
        for paragraph in _paragraphs(read["content"]["slides"][0]["shapes"])
        for run in paragraph["runs"]
        if run["kind"] == "text"
    )
    step = {"op": "replaceText", "target": anchor, "text": "Replaced"}
    assert deck.apply_edits({"expectVersion": read["version"], "steps": [step]})["ok"]
    story = next(
        story
        for story in deck.read_content()["stories"]
        if story["storyId"] == anchor["storyId"]
    )
    assert story["text"][anchor["start"] : anchor["start"] + 8] == "Replaced"


def test_bytes_exports_are_deterministic_snapshots(sample_bytes):
    content = bo.export_pptx_structured(sample_bytes, options={"includeComments": True})
    assert content["anchorScope"] == "snapshot"
    assert content == bo.export_pptx_structured(sample_bytes, options={"includeComments": True})
    provenance = content["slides"][0]["provenance"]
    assert provenance["part"] == "ppt/slides/slide1.xml"
    assert len(provenance["partSha256"]) == 64

    rendered = bo.render_pptx_markdown(content)
    assert rendered == bo.export_pptx_markdown(sample_bytes, options={"includeComments": True})
    assert rendered["markdown"].startswith("<!-- pptx-export:0 -->\n## Slide 1")

    truncated = bo.export_pptx_structured(sample_bytes, options={"maxBlocks": 5})
    assert truncated["truncated"]
    assert truncated["diagnostics"][-1]["code"] == "truncated"


def test_bytes_exports_raise_refusals_and_unreadable_input(sample_bytes):
    with pytest.raises(bo.ExportError) as refused:
        bo.export_pptx_structured(sample_bytes, options={"maxBlocks": 0})
    assert refused.value.failure == {
        "code": "invalid-options",
        "target": None,
        "message": "maxBlocks must be at least 1",
    }
    with pytest.raises(bo.ParseError):
        bo.export_pptx_structured(b"not a deck")
    with pytest.raises(ValueError):
        bo.export_pptx_structured(sample_bytes, options={"pages": True})
    content = bo.export_pptx_structured(sample_bytes)
    with pytest.raises(bo.ExportError) as limited:
        bo.render_pptx_markdown(content, max_bytes=10)
    assert limited.value.failure["code"] == "invalid-options"
    with pytest.raises(ValueError):
        bo.render_pptx_markdown({**content, "schemaVersion": 2})
