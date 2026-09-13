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
