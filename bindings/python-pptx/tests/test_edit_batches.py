import pytest
import betteroffice_pptx as bo


def _story(read):
    return next(story for story in read["stories"] if story["text"].split(" ")[0].isalpha())


def _within(story):
    return {key: story[key] for key in ("slideId", "shapeId", "storyId")}


def _range(story, start, end):
    return {"kind": "range", **_within(story), "start": start, "end": end}


def test_read_validate_apply_undo_and_save(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    read = deck.read_content()
    assert read["ok"] and read["version"] == deck.version()
    story = _story(read)
    request = {
        "expectVersion": read["version"],
        "steps": [{"op": "insertText", "target": _range(story, 0, 0), "at": "start", "text": "Draft: "}],
    }
    validation = deck.validate_edits(request)
    assert validation["ok"] and validation["wouldApply"]
    assert deck.version() == read["version"]
    assert not deck.is_edited

    applied = deck.apply_edits(request)
    assert applied["ok"] and applied["applied"] and applied["source"] == "host"
    assert applied["receipts"] == [
        {"stepIndex": 0, "changed": True, "target": {"kind": "range", **_range(story, 0, 7)}}
    ]
    assert applied["version"] == deck.version()
    assert deck.is_edited
    assert deck.story(story["storyId"]).text.startswith("Draft: ")

    stale = deck.apply_edits(request)
    assert stale["ok"] is False
    assert stale["failure"]["code"] == "stale-version"
    assert stale["version"] == deck.version()

    reopened = bo.Presentation.open(deck.save())
    assert reopened.story(story["storyId"]).text.startswith("Draft: ")
    assert reopened.version() != deck.version()
    assert deck.undo()
    assert deck.story(story["storyId"]).text == story["text"]


def test_refusals_no_ops_and_malformed_requests(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    read = deck.read_content()
    story = _story(read)
    refused = deck.apply_edits({
        "expectVersion": read["version"],
        "steps": [{
            "op": "deleteText",
            "target": {"kind": "search", "within": _within(story), "text": "no such text anywhere"},
        }],
    })
    assert refused["ok"] is False
    assert refused["failure"]["code"] == "missing-target"
    assert refused["failure"]["stepIndex"] == 0
    noop = deck.apply_edits({
        "expectVersion": read["version"],
        "steps": [{"op": "insertText", "target": _range(story, 0, 0), "at": "start", "text": ""}],
    })
    assert noop["ok"] and not noop["applied"]
    assert noop["version"] == read["version"]
    assert not deck.is_edited
    with pytest.raises(ValueError):
        deck.apply_edits({"expectVersion": read["version"], "steps": [{"op": "explode"}]})
    with pytest.raises(ValueError):
        deck.validate_edits({"steps": []})
    with pytest.raises(ValueError):
        deck.apply_edits({
            "expectVersion": read["version"],
            "steps": [{"op": "setShapeStroke", "target": {"slideId": story["slideId"], "shapeId": story["shapeId"]},
                       "stroke": {"widthPt": float("nan")}}],
        })
    assert deck.version() == read["version"]
    missing = deck.read_content({"slideIds": ["slide:missing"]})
    assert missing["ok"] is False and missing["failure"]["code"] == "missing-target"
    found = deck.find_text({"text": "e", "limit": 1})
    assert found["ok"] and found["truncated"] and len(found["matches"]) == 1


def test_agent_batches_can_stay_out_of_history(sample_bytes):
    deck = bo.Presentation.open_collaborative(sample_bytes)
    read = deck.read_content()
    story = _story(read)
    word = story["text"].split(" ")[0]
    applied = deck.apply_edits({
        "expectVersion": read["version"],
        "source": "agent",
        "history": "none",
        "steps": [{
            "op": "replaceText",
            "target": {"kind": "search", "within": _within(story), "text": word},
            "text": "Revised",
            "expect": {"text": word},
        }],
    })
    assert applied["ok"] and applied["applied"] and applied["source"] == "agent"
    assert applied["changedStories"] == [story["storyId"]]
    assert not deck.can_undo
    assert deck.story(story["storyId"]).text.startswith("Revised")
