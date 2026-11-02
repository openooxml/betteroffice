import inspect

import pytest

import betteroffice_pptx as bo


def _plain(value):
    if value is None or isinstance(value, (str, int, float, bool, bytes)):
        return value
    if isinstance(value, (list, tuple)):
        return [_plain(item) for item in value]
    if isinstance(value, dict):
        return {key: _plain(item) for key, item in value.items()}
    return {
        name: _plain(getattr(value, name))
        for name in dir(value)
        if not name.startswith("_") and not callable(getattr(value, name))
    }


@pytest.fixture
def presentations(sample_bytes):
    return tuple(
        bo.Presentation.open_collaborative(sample_bytes, client_id=101)
        for _ in range(2)
    )


def test_encoding_aliases_match_original_bytes(presentations):
    original, alias = presentations
    remote_state_vector = original.state_vector()
    for deck in presentations:
        deck.move_shape(0, deck.slide(0).shapes[0].id, 111, 222)
    assert _plain(alias.snapshot()) == _plain(original.snapshot())
    assert alias.encode_state_vector() == original.state_vector()
    assert alias.encode_state_as_update() == original.state_as_update()
    assert alias.encode_state_as_update(None) == original.state_as_update()
    expected = original.diff(remote_state_vector)
    assert alias.encode_diff(remote_state_vector) == expected
    assert alias.encode_state_as_update(remote_state_vector) == expected
    assert alias.encode_state_as_update(remote_state_vector=remote_state_vector) == expected
    current = original.state_vector()
    assert alias.encode_diff(current) == original.diff(current)
    assert alias.encode_state_as_update(current) == original.diff(current)
    before = _plain(alias.snapshot())
    for invalid in (b"", b"\xff"):
        with pytest.raises(bo.PptxError) as original_error:
            original.diff(invalid)
        for method in (alias.encode_diff, alias.encode_state_as_update):
            with pytest.raises(type(original_error.value)) as alias_error:
                method(invalid)
            assert str(alias_error.value) == str(original_error.value)
    assert _plain(alias.snapshot()) == before


@pytest.mark.parametrize("buffer_type", [bytes, bytearray, memoryview])
def test_public_encoding_aliases_accept_the_same_buffers(sample_bytes, buffer_type):
    original = bo.Presentation.open_collaborative(sample_bytes, client_id=101)
    alias = bo.Presentation.open_collaborative(sample_bytes, client_id=101)
    state_vector = buffer_type(original.state_vector())
    for deck in (original, alias):
        deck.move_shape(0, deck[0].shapes[0].id, 111, 222)
    assert alias.encode_diff(state_vector) == original.diff(state_vector)
    assert alias.encode_state_as_update(state_vector) == original.diff(state_vector)


def test_encoding_aliases_keep_the_collaborative_gate(sample_bytes):
    original = bo.Presentation.open(sample_bytes)
    alias = bo.Presentation.open(sample_bytes)
    for old_name, new_name, args in (
        ("state_vector", "encode_state_vector", ()),
        ("state_as_update", "encode_state_as_update", ()),
        ("diff", "encode_diff", (b"\x00",)),
        ("diff", "encode_state_as_update", (b"\x00",)),
    ):
        with pytest.raises(bo.NotCollaborativeError) as original_error:
            getattr(original, old_name)(*args)
        with pytest.raises(bo.NotCollaborativeError) as alias_error:
            getattr(alias, new_name)(*args)
        assert str(alias_error.value) == str(original_error.value)


def test_undo_boundary_alias_matches_original_grouping(presentations):
    original, alias = presentations
    before = _plain(original.snapshot())
    for deck, boundary in (
        (original, original.add_undo_barrier), (alias, alias.add_undo_boundary),
    ):
        deck.set_undo_capture_mode("manual")
        shape = deck.slide(0).shapes[0].id
        deck.move_shape(0, shape, 111, 111)
        deck.move_shape(0, shape, 222, 222)
        assert boundary() is None
        deck.move_shape(0, shape, 333, 333)
    for x in (222, before["slides"][0]["shapes"][0]["x"]):
        assert original.undo()
        assert alias.undo()
        assert alias.slide(0).shapes[0].x == x
        assert _plain(alias.snapshot()) == _plain(original.snapshot())
    assert _plain(alias.snapshot()) == before
    assert not alias.can_undo
    for _ in range(2):
        assert original.redo()
        assert alias.redo()
        assert _plain(alias.snapshot()) == _plain(original.snapshot())
    assert not alias.can_redo


def test_public_proposal_list_alias_matches_original(sample_bytes):
    original = bo.Presentation.open_collaborative(sample_bytes, client_id=101)
    alias = bo.Presentation.open_collaborative(sample_bytes, client_id=101)
    assert alias.list_proposals() == original.proposals() == []
    slide = original[0]
    shape = slide.shapes[0]
    edits = [{
        "type": "setShapeRect", "slideId": slide.id, "shapeId": shape.id,
        "rect": {"x": 111, "y": 222, "width": shape.width, "height": shape.height},
    }]
    proposal = original.propose("alias-test", edits, note="Move shape")
    assert alias.propose("alias-test", edits, note="Move shape") == proposal
    assert alias.list_proposals() == original.proposals() == [proposal]
    for deck in (original, alias):
        deck.move_shape(0, shape.id, 333, 444)
    assert alias.list_proposals() == original.proposals()
    assert alias.list_proposals()[0].stale_targets == (shape.id,)
    assert original.reject_proposal(proposal.id)
    assert alias.reject_proposal(proposal.id)
    assert alias.list_proposals() == original.proposals() == []


def test_public_alias_signatures_and_defaults():
    for name, parameters in (
        ("add_undo_boundary", ["self"]),
        ("list_proposals", ["self"]),
        ("encode_state_vector", ["self"]),
        ("encode_state_as_update", ["self", "remote_state_vector"]),
        ("encode_diff", ["self", "remote_state_vector"]),
    ):
        signature = inspect.signature(getattr(bo.Presentation, name))
        assert list(signature.parameters) == parameters
        for parameter in signature.parameters.values():
            assert parameter.kind is inspect.Parameter.POSITIONAL_OR_KEYWORD
            expected = (
                None if name == "encode_state_as_update" and parameter.name == "remote_state_vector"
                else inspect.Parameter.empty
            )
            assert parameter.default is expected
