import json
from dataclasses import FrozenInstanceError
from pathlib import Path

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


@pytest.fixture
def structural_proposal(sample_bytes, font_bytes):
    deck = bo.Presentation.open_collaborative(sample_bytes, client_id=9202)
    deck.register_font("Liberation Sans", font_bytes)
    live_ids = deck.slide_ids
    removed, _, moved = live_ids[:3]
    proposal = deck.propose("agent", [
        {"type": "insertSlide", "ref": "draft", "afterSlideId": None},
        {"type": "setSlideNotes", "slideId": "@draft", "text": "Draft notes"},
        {"type": "moveSlide", "slideId": moved, "afterSlideId": "@draft"},
        {"type": "deleteSlide", "slideId": removed},
        {"type": "addShape", "ref": "box", "slideId": "@draft", "draft": {
            "name": "Draft box", "geometry": "rect",
            "rect": {"x": 100000, "y": 100000, "width": 1000000, "height": 500000},
        }},
    ])
    return deck, proposal, live_ids


def test_proposal_slide_strip_preserves_preview_order_and_deleted_ghosts(structural_proposal):
    deck, proposal, live_ids = structural_proposal
    removed, unchanged, moved = live_ids[:3]
    inserted = proposal.refs["draft"]
    entries = deck.proposal_slide_strip(proposal.id)
    assert entries == [
        bo.ProposalStripEntry(removed, "deleted", 0, None),
        bo.ProposalStripEntry(inserted, "inserted", None, 0),
        bo.ProposalStripEntry(moved, "moved", 2, 1),
        bo.ProposalStripEntry(unchanged, "unchanged", 1, 2),
        *[
            bo.ProposalStripEntry(slide_id, "unchanged", index, index)
            for index, slide_id in enumerate(live_ids[3:], start=3)
        ],
    ]
    assert [entry.slide_id for entry in entries if entry.preview_index is not None] == [
        slide["id"] for slide in deck.preview_proposal(proposal.id).snapshot["slides"]
    ]
    assert "ProposalStripEntry" in bo.__all__
    with pytest.raises(FrozenInstanceError):
        entries[0].status = "unchanged"
    assert deck.slide_ids == live_ids
    assert not deck.can_undo


def test_proposal_structural_changes_and_refs_are_populated(structural_proposal):
    deck, proposal, live_ids = structural_proposal
    changes = {change.slide_id: change for change in proposal.changes if change.shape_id is None}
    for slide_id, kind, before, after in (
        (proposal.refs["draft"], "inserted", None, 0),
        (live_ids[2], "moved", 2, 1),
        (live_ids[0], "deleted", 0, None),
    ):
        change = changes[slide_id]
        assert (change.kind, change.before_index, change.after_index) == (kind, before, after)
        assert change.before_index_present
        assert change.after_index_present
    shape_change = next(change for change in proposal.changes if change.shape_id == proposal.refs["box"])
    assert shape_change.kind == "inserted"
    assert shape_change.before_index is None
    assert shape_change.after_index is None
    assert not shape_change.before_index_present
    assert not shape_change.after_index_present
    assert deck.proposals() == [proposal]
    assert deck.preview_proposal(proposal.id).proposal == proposal


def test_proposal_diff_by_id_renders_inserted_existing_and_deleted_slides(structural_proposal):
    deck, proposal, live_ids = structural_proposal
    inserted = deck.layout_proposal_diff_slide_by_id(proposal.id, proposal.refs["draft"])
    assert inserted["snapshot"]["slides"][0]["id"] == proposal.refs["draft"]
    assert inserted["snapshot"]["slides"][0]["notes"] == "Draft notes"
    assert inserted["proposal"]["refs"] == proposal.refs
    assert inserted["frame"]["primitives"]
    assert inserted["shapeChanges"] == [{"shapeId": proposal.refs["box"], "kind": "inserted"}]
    assert inserted["textChanges"] == []
    for slide_id in (live_ids[2], live_ids[0]):
        diff = deck.layout_proposal_diff_slide_by_id(proposal.id, slide_id)
        assert diff == deck.layout_proposal_diff_slide(proposal.id, slide_id)
        assert diff["snapshot"]["slides"][0]["id"] == slide_id
        assert diff["frame"]["primitives"]
        assert "shapeChanges" not in diff
    assert deck.slide_ids == live_ids
    assert not deck.can_undo


def test_proposal_diff_by_id_retains_deleted_shape_geometry(font_bytes):
    path = Path(__file__).resolve().parents[3] / "crates/pptx-edit/tests/fixtures/proposal-shapes.pptx"
    deck = bo.Presentation.open_collaborative(path.read_bytes(), client_id=9203)
    deck.register_font("Liberation Sans", font_bytes)
    slide = deck[0]
    shape = slide.shapes[0]
    proposal = deck.propose("agent", [
        {"type": "addTextBox", "ref": "box", "slideId": slide.id, "draft": {
            "name": "New box", "text": "", "style": {},
            "rect": {"x": 100000, "y": 100000, "width": 1000000, "height": 500000},
        }},
        {"type": "replaceText", "storyId": "@box", "start": 0, "end": 0, "text": "Filled box"},
        {"type": "removeShape", "slideId": slide.id, "shapeId": shape.id},
    ])
    diff = deck.layout_proposal_diff_slide_by_id(proposal.id, slide.id)
    assert {"shapeId": proposal.refs["box"], "kind": "inserted"} in diff["shapeChanges"]
    assert {"shapeId": shape.id, "kind": "deleted"} in diff["shapeChanges"]
    assert any(primitive.get("shapeId") == shape.id for primitive in diff["frame"]["primitives"])
    assert [value.id for value in deck[0].shapes] == [value.id for value in slide.shapes]


def test_render_proposal_resolves_inserted_ids_and_preview_indices(structural_proposal):
    deck, proposal, live_ids = structural_proposal
    inserted = deck.render_proposal(proposal.id, proposal.refs["draft"]).to_dict()
    assert inserted == deck.render_proposal(proposal.id, 0).to_dict()
    assert any(primitive.get("shapeId") == proposal.refs["box"] for primitive in inserted["primitives"])
    moved = deck.render_proposal(proposal.id, live_ids[2]).to_dict()
    assert moved == deck.render_proposal(proposal.id, 1).to_dict()
    assert deck.slide_ids == live_ids


def test_render_proposal_accepts_inserted_index_past_live_deck(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    live_count = len(deck)
    proposal = deck.propose("agent", [
        {"type": "insertSlide", "ref": "tail", "afterSlideId": deck.slide_ids[-1]},
        {"type": "addShape", "slideId": "@tail", "draft": {
            "name": "Tail box", "geometry": "rect",
            "rect": {"x": 100000, "y": 100000, "width": 1000000, "height": 500000},
        }},
    ])
    frame = deck.render_proposal(proposal.id, live_count)
    assert len(frame) > 0
    assert frame.to_dict() == deck.render_proposal(proposal.id, proposal.refs["tail"]).to_dict()
    assert len(deck) == live_count


def test_proposal_preview_errors_match_existing_diff_errors(structural_proposal):
    deck, proposal, live_ids = structural_proposal
    with pytest.raises(KeyError) as missing_proposal:
        deck.layout_proposal_diff_slide("missing-proposal", live_ids[0])
    for read in (
        lambda: deck.proposal_slide_strip("missing-proposal"),
        lambda: deck.layout_proposal_diff_slide_by_id("missing-proposal", live_ids[0]),
        lambda: deck.render_proposal("missing-proposal", live_ids[0]),
    ):
        with pytest.raises(type(missing_proposal.value)) as error:
            read()
        assert error.value.args == missing_proposal.value.args
    with pytest.raises(KeyError) as missing_slide:
        deck.layout_proposal_diff_slide(proposal.id, "missing-slide")
    for read in (
        lambda: deck.layout_proposal_diff_slide_by_id(proposal.id, "missing-slide"),
        lambda: deck.render_proposal(proposal.id, "missing-slide"),
        lambda: deck.render_proposal(proposal.id, live_ids[0]),
    ):
        with pytest.raises(type(missing_slide.value)):
            read()


@pytest.mark.parametrize("slide, error", [(True, TypeError), (-1, IndexError), (1.5, TypeError), (10000, IndexError)])
def test_render_proposal_preserves_invalid_target_errors(structural_proposal, slide, error):
    deck, proposal, _ = structural_proposal
    with pytest.raises(error):
        deck.render_proposal(proposal.id, slide)


def test_proposal_dataclasses_preserve_optional_fields_and_positional_constructors():
    original = bo.ProposalChange("slide", None, None, None, "old", "new")
    omitted = {
        "slideId": "slide", "shapeId": None, "before": None, "after": None,
        "oldText": "old", "newText": "new",
    }
    assert bo.ProposalChange._from_dict(omitted) == original
    explicit_null = bo.ProposalChange._from_dict({**omitted, "beforeIndex": None, "afterIndex": None})
    assert explicit_null.before_index is original.before_index is None
    assert explicit_null.after_index is original.after_index is None
    assert explicit_null.before_index_present and explicit_null.after_index_present
    assert not original.before_index_present and not original.after_index_present
    assert explicit_null != original
    for key, before_present, after_present in (
        ("beforeIndex", True, False), ("afterIndex", False, True),
    ):
        change = bo.ProposalChange._from_dict({**omitted, key: None})
        assert change.before_index_present is before_present
        assert change.after_index_present is after_present
    proposal = bo.Proposal("id", "agent", None, (), (original,), ())
    wire = {"id": "id", "agentId": "agent", "note": None, "edits": [], "changes": [omitted], "staleTargets": []}
    assert bo.Proposal._from_dict(wire) == proposal
    assert bo.Proposal._from_dict({**wire, "refs": {}}) == proposal
    assert bo.Proposal._from_dict({**wire, "refs": {"draft": "slide"}}).refs == {"draft": "slide"}
    proposal.refs["local"] = "slide"
    assert bo.Proposal("other", "agent", None, (), (), ()).refs == {}


def test_nonstructural_proposal_omits_indices_and_empty_refs(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    proposal = deck.propose("agent", [{"type": "setSlideNotes", "slideId": deck[0].id, "text": "Notes"}])
    change = proposal.changes[0]
    assert change.kind is None
    assert change.before_index is None
    assert change.after_index is None
    assert not change.before_index_present
    assert not change.after_index_present
    assert proposal.refs == {}
    assert deck.proposal_slide_strip(proposal.id)[0] == bo.ProposalStripEntry(deck[0].id, "changed", 0, 0)
    wire = json.loads(deck._inner.proposals_json())[0]
    assert "refs" not in wire
    assert "beforeIndex" not in wire["changes"][0]
    assert "afterIndex" not in wire["changes"][0]
