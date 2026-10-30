import gc
import io
import threading
import weakref
import zipfile

import pytest

import betteroffice_pptx as bo
from betteroffice_pptx import _betteroffice_pptx as native


def _plain(value):
    if value is None or isinstance(value, (str, int, float, bool, bytes)):
        return value
    if isinstance(value, (list, tuple)):
        return [_plain(item) for item in value]
    if isinstance(value, dict):
        return {key: _plain(item) for key, item in value.items()}
    result = {}
    for name in dir(value):
        if not name.startswith("_"):
            item = getattr(value, name)
            if not callable(item):
                result[name] = _plain(item)
    return result


def _table_source(sample_bytes):
    cell = (
        '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Cell</a:t>'
        '</a:r></a:p></a:txBody><a:tcPr/></a:tc>'
    )
    row = '<a:tr h="400000">' + cell * 2 + '</a:tr>'
    table = (
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="10000" name="Collaboration table"/>'
        '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm>'
        '<a:off x="0" y="0"/><a:ext cx="2000000" cy="800000"/></p:xfrm><a:graphic>'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
        '<a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="1000000"/><a:gridCol w="1000000"/>'
        '</a:tblGrid>' + row * 2 + '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
    )
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(sample_bytes)) as source:
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
            for part in source.infolist():
                data = source.read(part.filename)
                if part.filename == "ppt/slides/slide1.xml":
                    data = data.replace(b"</p:spTree>", table.encode() + b"</p:spTree>")
                target.writestr(part, data)
    return output.getvalue()


def test_peers_exchange_updates_with_equal_snapshots(sample_bytes):
    left = bo.Presentation.open_collaborative(sample_bytes, client_id=101)
    right = bo.Presentation.open_collaborative(sample_bytes, client_id=202)
    shape = left[0].shapes[0].id
    left.move_shape(0, shape, 111, 222)
    merged = right.apply_update(left.diff(right.state_vector()))
    assert _plain(merged) == _plain(left.snapshot())
    right.move_shape(0, shape, 333, 444)
    left.apply_update(right.diff(left.state_vector()))
    assert _plain(left.snapshot()) == _plain(right.snapshot())


@pytest.mark.parametrize("buffer_type", [bytes, bytearray, memoryview])
def test_open_from_update_attaches_matching_source(sample_bytes, buffer_type):
    source = bo.Presentation.open_collaborative(sample_bytes, client_id=101)
    shape = source[0].shapes[0].id
    source.move_shape(0, shape, 111, 222)
    joined = bo.Presentation.open_from_update(
        buffer_type(source.state_as_update()), client_id=202,
        source=buffer_type(sample_bytes),
    )
    assert joined.client_id == 202
    assert joined.is_collaborative
    assert _plain(joined.snapshot()) == _plain(source.snapshot())
    saved = bo.Presentation.open(joined.save())
    assert (saved[0].shapes[0].x, saved[0].shapes[0].y) == (111, 222)
    joined.move_shape(0, shape, 333, 444)
    source.apply_update(joined.diff(source.state_vector()))
    assert _plain(joined.snapshot()) == _plain(source.snapshot())


@pytest.mark.parametrize("source_kind", ["absent", "different", "invalid"])
def test_open_from_update_with_foreign_source_opens_but_cannot_save(sample_bytes, source_kind):
    source = bo.Presentation.open_collaborative(sample_bytes, client_id=101)
    source.move_shape(0, source[0].shapes[0].id, 111, 222)
    data = {"absent": None, "different": source.save(), "invalid": b"foreign"}[source_kind]
    joined = bo.Presentation.open_from_update(
        source.state_as_update(), client_id=202, source=data,
    )
    assert joined.slide_count == source.slide_count
    assert (joined[0].shapes[0].x, joined[0].shapes[0].y) == (111, 222)
    with pytest.raises(bo.PptxError):
        joined.save()
    joined.move_shape(0, joined[0].shapes[0].id, 333, 444)
    source.apply_update(joined.diff(source.state_vector()))
    assert source[0].shapes[0].x == 333


def test_open_from_update_rejects_invalid_updates_and_client_ids(sample_bytes):
    source = bo.Presentation.open_collaborative(sample_bytes, client_id=101)
    update = source.state_as_update()
    for client_id in (0, 2**60):
        with pytest.raises(ValueError, match="client ID"):
            bo.Presentation.open_from_update(update, client_id=client_id)
    with pytest.raises(bo.InvalidUpdateError):
        bo.Presentation.open_from_update(b"\xff", client_id=202)
    with pytest.raises(TypeError):
        bo.Presentation.open_from_update(update)


def test_marking_collaborative_refuses_table_structure_edits(sample_bytes):
    deck = bo.Presentation.open(_table_source(sample_bytes))
    table = next(shape for shape in deck[0].shapes if shape.name == "Collaboration table")
    assert deck.is_collaborative is False
    deck.insert_table_row(0, table.id, 0, "after")
    assert deck.can_undo
    deck.set_collaborative()
    assert deck.is_collaborative is True
    assert deck.can_undo is False
    before = _plain(deck.snapshot())
    for method in (
        "insert_table_row", "delete_table_row", "insert_table_column", "delete_table_column",
    ):
        with pytest.raises(bo.PptxError, match="collaborative mode"):
            deck.check_table_edit(method, 0, table.id, 0)
        args = (0, table.id, 0, "before") if method.startswith("insert") else (0, table.id, 0)
        with pytest.raises(bo.PptxError, match="collaborative mode"):
            getattr(deck, method)(*args)
    assert _plain(deck.snapshot()) == before
    with pytest.raises(bo.NotCollaborativeError):
        deck.state_vector()


def test_manual_undo_capture_groups_edits_until_a_barrier(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    before = _plain(deck.snapshot())
    shape = deck[0].shapes[0].id
    assert deck.undo_capture_mode == "auto"
    deck.set_undo_capture_mode("manual")
    assert deck.undo_capture_mode == "manual"
    deck.move_shape(0, shape, 111, 111)
    deck.move_shape(0, shape, 222, 222)
    deck.add_undo_barrier()
    deck.move_shape(0, shape, 333, 333)
    assert deck.undo()
    assert deck[0].shapes[0].x == 222
    assert deck.undo()
    assert _plain(deck.snapshot()) == before
    assert not deck.can_undo
    assert deck.redo()
    assert deck[0].shapes[0].x == 222
    deck.set_undo_capture_mode("auto")
    with pytest.raises(ValueError, match="auto or manual"):
        deck.set_undo_capture_mode("invalid")
    assert deck.undo_capture_mode == "auto"


def test_updates_are_delivered_on_the_owner_after_the_transaction(sample_bytes):
    left = bo.Presentation.open_collaborative(sample_bytes, client_id=101)
    right = bo.Presentation.open_collaborative(sample_bytes, client_id=202)
    owner = threading.get_ident()
    events = []

    def receive(deck, update, origin):
        assert threading.get_ident() == owner
        assert isinstance(update, bytes)
        assert deck[0].shapes[0].x == 111
        assert deck.snapshot().slides
        events.append((update, origin))

    with left.on_update(lambda update, origin: receive(left, update, origin)) as local:
        assert isinstance(local, bo.UpdateSubscription)
        with right.on_update(lambda update, origin: receive(right, update, origin)):
            left.move_shape(0, left[0].shapes[0].id, 111, 222)
            assert len(events) == 1
            assert events[0][1] == "local"
            right.apply_update(events[0][0])
            assert len(events) == 2
            assert events[1][1] == "remote"
            assert _plain(left.snapshot()) == _plain(right.snapshot())
    left.move_shape(0, left[0].shapes[0].id, 333, 444)
    right.apply_update(left.diff(right.state_vector()))
    assert len(events) == 2


def test_unsubscribe_and_drop_stop_updates(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    events = []
    subscription = deck.on_update(lambda update, origin: events.append((update, origin)))
    shape = deck[0].shapes[0].id
    deck.move_shape(0, shape, 111, 222)
    assert len(events) == 1
    subscription.unsubscribe()
    subscription.unsubscribe()
    deck.move_shape(0, shape, 333, 444)
    assert len(events) == 1
    subscription = deck.on_update(lambda update, origin: events.append((update, origin)))
    del subscription
    gc.collect()
    deck.move_shape(0, shape, 555, 666)
    assert len(events) == 1
    with pytest.raises(TypeError, match="callable"):
        deck.on_update(None)


def test_callback_errors_are_reraised_after_all_listeners_and_leave_the_edit_committed(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    events = []
    failures = []

    def fail(update, origin):
        failures.append((update, origin))
        raise RuntimeError("update callback failed")

    with deck.on_update(fail):
        with deck.on_update(lambda update, origin: events.append((update, origin))):
            with pytest.raises(RuntimeError, match="update callback failed"):
                deck.move_shape(0, deck[0].shapes[0].id, 111, 222)
            assert deck.is_edited
            assert deck[0].shapes[0].x == 111
            with pytest.raises(RuntimeError, match="update callback failed"):
                deck.move_shape(0, deck[0].shapes[0].id, 333, 444)
    assert events == failures
    assert len(events) == 2
    assert deck.is_edited
    assert deck[0].shapes[0].x == 333
    assert bo.Presentation.open(deck.save())[0].shapes[0].x == 333
    assert deck.undo()
    assert deck.redo()
    deck.move_shape(0, deck[0].shapes[0].id, 555, 666)
    assert len(events) == 2


def test_subscription_callback_cycles_are_collectable_on_the_owner(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    events = []

    def subscribe():
        holder = {}

        def receive(update, origin):
            events.append((update, origin))
            holder["subscription"].unsubscribe()

        holder["subscription"] = deck.on_update(receive)
        return weakref.ref(receive)

    callback = subscribe()
    gc.collect()
    assert callback() is None
    deck.move_shape(0, deck[0].shapes[0].id, 111, 222)
    assert events == []


def test_raising_remote_callback_leaves_the_applied_update_committed(sample_bytes):
    left = bo.Presentation.open_collaborative(sample_bytes, client_id=101)
    right = bo.Presentation.open_collaborative(sample_bytes, client_id=202)
    events = []

    def receive(update, origin):
        events.append(origin)
        assert origin == "remote"
        assert right[0].shapes[0].x == 111
        raise ValueError("remote callback failed")

    left.move_shape(0, left[0].shapes[0].id, 111, 222)
    with right.on_update(receive):
        with pytest.raises(ValueError, match="remote callback failed"):
            right.apply_update(left.diff(right.state_vector()))
    assert events == ["remote"]
    assert right.is_edited
    assert _plain(right.snapshot()) == _plain(left.snapshot())
    assert bo.Presentation.open(right.save())[0].shapes[0].x == 111


def test_reentrant_edits_are_queued_and_unsubscribe_inside_callback_is_safe(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    shape = deck[0].shapes[0].id
    positions = []

    def receive(update, origin):
        assert origin == "local"
        positions.append(deck[0].shapes[0].x)
        if len(positions) == 1:
            deck.move_shape(0, shape, 333, 444)
        else:
            subscription.unsubscribe()

    subscription = deck.on_update(receive)
    deck.move_shape(0, shape, 111, 222)
    assert positions == [111, 333]
    deck.move_shape(0, shape, 555, 666)
    assert positions == [111, 333]


def test_batch_proposal_and_history_calls_flush_updates(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    events = []
    shape = deck[0].shapes[0]
    slide = deck[0].id
    with deck.on_update(lambda update, origin: events.append((update, origin))):
        applied = deck.apply_edits({
            "expectVersion": deck.version(), "source": "agent", "history": "separate",
            "steps": [{
                "op": "setShapeRect", "target": {"slideId": slide, "shapeId": shape.id},
                "rect": {"x": 111, "y": 222, "width": shape.width, "height": shape.height},
            }],
        })
        assert applied["ok"]
        assert events and all(origin == "local" for _, origin in events)
        count = len(events)
        assert deck.undo()
        assert len(events) > count
        count = len(events)
        assert deck.redo()
        assert len(events) > count
        count = len(events)
        assert deck.history_step("undo")["applied"]
        assert len(events) > count
        proposal = deck.propose("collab-test", [{
            "type": "setShapeRect", "slideId": slide, "shapeId": shape.id,
            "rect": {"x": 333, "y": 444, "width": shape.width, "height": shape.height},
        }])
        count = len(events)
        assert deck.accept_proposal(proposal.id)
        assert len(events) > count


def test_native_binding_exposes_the_same_collaboration_surface(sample_bytes):
    source = native.Presentation.open_collaborative(sample_bytes, client_id=101)
    deck = native.Presentation.open_from_update(
        source.state_as_update(), client_id=202, source=sample_bytes,
    )
    deck.set_collaborative()
    assert deck.is_collaborative
    deck.set_undo_capture_mode("manual")
    assert deck.undo_capture_mode == "manual"
    events = []
    with deck.on_update(lambda update, origin: events.append((update, origin))):
        deck.move_shape(0, deck.slide(0).shapes[0].id, 111, 222)
    assert len(events) == 1
    assert events[0][1] == "local"
    assert deck.save().startswith(b"PK")
