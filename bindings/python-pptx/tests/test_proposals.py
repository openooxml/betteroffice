import pytest
import betteroffice_pptx as bo


def test_stage_preview_accept_undo_and_save(sample_bytes, font_bytes):
    deck = bo.Presentation.open(sample_bytes)
    deck.register_font("Liberation Sans", font_bytes)
    shape = next(shape for shape in deck[0].shapes if shape.stories)
    story = shape.stories[0]
    old_text = story.text
    end = len(old_text.split("\n")[0].encode("utf-16-le")) // 2
    proposal = deck.propose("python-agent", [{
        "type": "replaceText", "storyId": story.id, "start": 0, "end": end, "text": "A reviewed title",
    }], note="Tighten the title")
    assert isinstance(proposal, bo.Proposal)
    assert proposal.agent_id == "python-agent"
    assert "A reviewed title" in proposal.changes[0].new_text
    assert deck.story(story.id).text == old_text
    assert not deck.can_undo
    assert deck.preview_proposal(proposal.id).proposal.id == proposal.id
    assert len(deck.render_proposal(proposal.id, 0)) > 0
    assert deck.accept_proposal(proposal.id)
    assert "A reviewed title" in deck.story(story.id).text
    reopened = bo.Presentation.open(deck.save())
    assert "A reviewed title" in next(shape for shape in reopened[0].shapes if shape.stories).stories[0].text
    assert reopened.proposals() == []
    assert deck.undo()
    assert deck.story(story.id).text == old_text
    assert deck.redo()
    assert "A reviewed title" in deck.story(story.id).text


def test_stale_force_reject_and_invalid_proposals(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    slide_id = deck[0].id
    shape = deck[0].shapes[0]
    edit = {
        "type": "setShapeRect", "slideId": slide_id, "shapeId": shape.id,
        "rect": {"x": shape.x + 100_000, "y": shape.y, "width": shape.width, "height": shape.height},
    }
    proposal = deck.propose("python-agent", [edit])
    deck.move_shape(slide_id, shape.id, shape.x + 200_000, shape.y)
    with pytest.raises(bo.StaleProposalError) as stale:
        deck.accept_proposal(proposal.id)
    assert stale.value.targets == [shape.id]
    assert deck.proposals()[0].stale_targets == (shape.id,)
    assert deck.preview_proposal(proposal.id).proposal.changes[0].before["x"] == shape.x + 200_000
    assert deck.accept_proposal(proposal.id, force=True)
    assert deck[0].shapes[0].x == shape.x + 100_000
    rejected = deck.propose("python-agent", [edit])
    assert deck.reject_proposal(rejected.id)
    assert not deck.reject_proposal(rejected.id)
    assert deck[0].shapes[0].x == shape.x + 100_000
    with pytest.raises(KeyError):
        deck.accept_proposal(rejected.id)
    with pytest.raises(ValueError):
        deck.propose("python-agent", [{"type": "setSlideNotes", "slideId": slide_id, "text": "Bad", "typo": True}])


@pytest.mark.parametrize("kind", ["insertTableRow", "deleteTableRow", "insertTableColumn", "deleteTableColumn"])
def test_table_proposal_dicts_preview_accept_reject_undo_and_stale(kind, font_bytes):
    import io
    import re
    import zipfile
    from pathlib import Path

    fixture = Path(__file__).resolve().parents[3] / "crates/pptx-render/tests/fixtures/table-basic.pptx"
    output = io.BytesIO()
    with zipfile.ZipFile(fixture) as source, zipfile.ZipFile(output, "w") as target:
        for part in source.infolist():
            data = source.read(part.filename)
            if part.filename == "ppt/slides/slide1.xml":
                data = re.sub(rb' (gridSpan|rowSpan|hMerge|vMerge)="[^"]*"', b"", data)
            target.writestr(part, data)
    deck = bo.Presentation.open(output.getvalue())
    deck.register_font("Liberation Sans", font_bytes)
    before = deck.read_content()["slides"]
    slide_id, shape_id = deck[0].id, deck[0].shapes[0].id
    edit = {"type": kind, "slideId": slide_id, "shapeId": shape_id, "index": 0}
    if kind.startswith("insert"):
        edit.update(position="before", cells=["Python cell"])
    if kind == "insertTableColumn":
        edit["width"] = "copy"
    proposal = deck.propose("python-agent", [edit])
    assert proposal.changes[0].shape_id == shape_id
    preview = deck.preview_proposal(proposal.id)
    table = preview.snapshot["slides"][0]["shapes"][0]["graphic"]
    rows = 4 if kind == "insertTableRow" else 2 if kind == "deleteTableRow" else 3
    columns = 4 if kind == "insertTableColumn" else 2 if kind == "deleteTableColumn" else 3
    assert (len(table["rows"]), len(table["grid"])) == (rows, columns)
    assert len(deck.render_proposal(proposal.id, 0)) > 0
    assert deck.read_content()["slides"] == before and not deck.can_undo
    assert deck.reject_proposal(proposal.id)
    assert deck.read_content()["slides"] == before
    proposal = deck.propose("python-agent", [edit])
    assert deck.accept_proposal(proposal.id)
    reopened = bo.Presentation.open(deck.save())
    saved_table = reopened.read_content()["slides"][0]["shapes"][0]["graphic"]
    assert (len(saved_table["rows"]), len(saved_table["grid"])) == (rows, columns)
    assert deck.undo()
    assert deck.read_content()["slides"] == before and not deck.can_undo
    for invalid in [dict(edit, index=99), dict(edit, extra=True)]:
        with pytest.raises((bo.PptxError, ValueError)):
            deck.propose("python-agent", [invalid])
        assert deck.read_content()["slides"] == before and deck.proposals() == []
    pending = deck.propose("python-agent", [edit])
    deck.insert_table_row(slide_id, shape_id, 0, "after")
    with pytest.raises(bo.StaleProposalError) as stale:
        deck.accept_proposal(pending.id)
    assert stale.value.targets == [shape_id]
