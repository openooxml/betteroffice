import gc
import os
import subprocess
import sys
import threading

import pytest

import betteroffice_pptx as bo

PPTX_MAGIC = b"PK\x03\x04"


@pytest.fixture
def deck(sample_bytes):
    return bo.Presentation.open(sample_bytes)


def test_open_exposes_slides(deck):
    assert deck.slide_count == len(deck) == 3
    assert len(deck.slide_ids) == 3
    assert [slide.index for slide in deck] == [0, 1, 2]
    assert deck.width_emu == 12192000 and deck.height_emu == 6858000


def test_open_path(sample_path):
    assert bo.Presentation.open_path(sample_path).slide_count == 3


def test_slide_lookup_by_index_and_id(deck):
    first = deck[0]
    assert deck.slide(first.id).index == 0
    assert deck[first.id].id == first.id

    snapshot = deck.snapshot()
    assert [slide.id for slide in snapshot.slides] == deck.slide_ids
    assert [slide.index for slide in snapshot.slides] == [0, 1, 2]
    assert len(snapshot) == 3


def test_shapes_carry_geometry_and_resolved_paint(deck):
    rail = deck[0].shapes[0]
    assert rail.name == "Cobalt rail"
    assert rail.kind == "shape"
    assert rail.geometry == "rect"
    assert (rail.x, rail.y, rail.width, rail.height) == (0, 0, 228600, 6858000)
    assert rail.fill_color == "#315EFB"

    card = next(s for s in deck[1].shapes if s.name == "DOCX card")
    assert card.geometry == "roundRect"
    assert card.outline_color == "#C7D1FF"
    assert card.adjustments


def test_shape_kinds_and_nesting(deck):
    kinds = {shape.kind for slide in deck for shape in slide.shapes}
    assert {"shape", "picture", "group", "graphicFrame"} <= kinds

    group = next(s for s in deck[0].shapes if s.kind == "group")
    assert len(group.children) == 14

    picture = next(s for s in deck[0].shapes if s.kind == "picture")
    assert picture.media_path == "ppt/media/betteroffice-mark.png"


def test_text_reads_through_stories(deck):
    title = next(s for s in deck[0].shapes if s.name == "Title")
    assert title.text == "Office files,\nwithout the office."
    assert [p.text for p in title.stories[0].paragraphs] == [
        "Office files,",
        "without the office.",
    ]
    assert "Office files," in deck[0].text
    assert len(deck[0]) == len(deck[0].shapes)


def test_story_reads_by_id(deck):
    story = next(s for s in deck[0].shapes if s.name == "Title").stories[0]
    assert deck.story(story.id).text == story.text
    assert deck.story(story.id).length == story.length


def test_media_parts_are_readable(deck, tmp_path):
    part = deck.media()[0]
    assert part.path == "ppt/media/betteroffice-mark.png"
    assert part.content_type == "image/png"
    assert part.bytes[:4] == b"\x89PNG"
    assert len(part) == len(part.bytes)

    target = tmp_path / "mark.png"
    part.write(target)
    assert target.read_bytes() == part.bytes


def test_layouts_are_the_paths_insert_slide_takes(deck):
    assert deck.layouts == ["ppt/slideLayouts/slideLayout1.xml"]
    edit = deck.insert_slide(1, layout=deck.layouts[0])
    assert deck.slide(edit.slide_id).layout == deck.layouts[0]


def test_move_shape_reports_before_and_after(deck):
    shape = deck[0].shapes[0]
    edit = deck.move_shape(0, shape.id, 111, 222)

    assert (edit.before.x, edit.before.y) == (shape.x, shape.y)
    assert (edit.after.x, edit.after.y) == (111, 222)
    assert (deck[0].shapes[0].x, deck[0].shapes[0].y) == (111, 222)


def test_resize_shape(deck):
    shape = deck[0].shapes[0]
    edit = deck.resize_shape(0, shape.id, 4000, 5000)

    assert (edit.after.width, edit.after.height) == (4000, 5000)
    assert (deck[0].shapes[0].width, deck[0].shapes[0].height) == (4000, 5000)


def test_add_text_box_creates_an_editable_story(deck):
    before = len(deck[0].shapes)
    edit = deck.add_text_box(
        0, x=100, y=200, width=1000, height=500, text="Hello", bold=True
    )

    assert len(deck[0].shapes) == before + 1
    added = deck[0].shapes[edit.index]
    assert added.id == edit.shape_id
    assert (added.x, added.y, added.width, added.height) == (100, 200, 1000, 500)
    assert added.text == "Hello"
    assert added.stories[0].paragraphs[0].runs[0].bold is True


def test_add_shape_applies_preset_defaults(deck):
    edit = deck.add_shape(
        0, "roundRect", x=0, y=0, width=1000, height=1000, fill="#123456"
    )
    added = deck[0].shapes[edit.index]

    assert added.geometry == "roundRect"
    assert added.fill_color == "#123456"
    assert "adj" in added.adjustments


def test_remove_shape(deck):
    edit = deck.add_shape(0, "rect", x=0, y=0, width=10, height=10)
    deck.remove_shape(0, edit.shape_id)
    assert edit.shape_id not in [shape.id for shape in deck[0].shapes]


def test_set_shape_fill_and_stroke(deck):
    edit = deck.add_shape(0, "rect", x=0, y=0, width=10, height=10, fill="#123456")

    fill = deck.set_shape_fill(0, edit.shape_id, "#654321")
    assert (fill.before, fill.after) == ("#123456", "#654321")
    assert deck[0].shapes[edit.index].fill_color == "#654321"

    stroke = deck.set_shape_stroke(0, edit.shape_id, color="#000000", width_pt=3.0)
    assert stroke.before is None
    assert stroke.after.color == "#000000"
    assert stroke.after.width_pt == pytest.approx(3.0)
    assert deck[0].shapes[edit.index].outline_color == "#000000"


def test_adjustments_are_clamped_not_rejected(deck):
    edit = deck.add_shape(0, "roundRect", x=0, y=0, width=1000, height=1000)
    adjust = deck.set_shape_adjust(0, edit.shape_id, {"adj": 0.9})

    assert adjust.after["adj"] == pytest.approx(0.5)
    assert deck[0].shapes[edit.index].adjustments["adj"] == pytest.approx(0.5)


def test_slide_lifecycle(deck):
    inserted = deck.insert_slide(1)
    assert deck.slide_count == 4
    assert deck.slide_ids[1] == inserted.slide_id
    assert inserted.from_index is None and inserted.to_index == 1

    moved = deck.move_slide(inserted.slide_id, 3)
    assert (moved.from_index, moved.to_index) == (1, 3)
    assert deck.slide_ids[3] == inserted.slide_id

    deleted = deck.delete_slide(inserted.slide_id)
    assert deleted.to_index is None
    assert deck.slide_count == 3
    assert inserted.slide_id not in deck.slide_ids


def test_insert_and_delete_text(deck):
    edit = deck.add_text_box(0, x=0, y=0, width=100, height=100, text="Hi")
    story = deck[0].shapes[edit.index].stories[0]

    inserted = deck.insert_text(story.id, 0, "Oh ")
    assert (inserted.start, inserted.end) == (0, 3)
    assert deck.story(story.id).text == "Oh Hi"

    removed = deck.delete_text(story.id, 0, 3)
    assert removed.text == "Oh "
    assert len(removed) == 3
    assert deck.story(story.id).text == "Hi"


def test_format_text_patches_only_what_is_passed(deck):
    edit = deck.add_text_box(0, x=0, y=0, width=100, height=100, text="Hi", bold=True)
    story = deck[0].shapes[edit.index].stories[0]

    deck.format_text(story.id, 0, 2, color="#ff0000")
    run = deck.story(story.id).paragraphs[0].runs[0]
    assert run.color == "#ff0000"
    assert run.bold is True


def test_insert_paragraph_break_splits_a_paragraph(deck):
    edit = deck.add_text_box(0, x=0, y=0, width=100, height=100, text="Hi")
    story = deck[0].shapes[edit.index].stories[0]
    assert len(deck.story(story.id).paragraphs) == 1

    deck.insert_paragraph_break(story.id, 1)
    assert [p.text for p in deck.story(story.id).paragraphs] == ["H", "i"]
    assert deck.story(story.id).text == "H\ni"


def test_delete_text_may_not_cross_a_paragraph_boundary(deck):
    story = next(s for s in deck[0].shapes if s.name == "Title").stories[0]
    assert len(story.paragraphs) == 2

    with pytest.raises(bo.RangeError, match="paragraph boundary"):
        deck.delete_text(story.id, 10, 20)
    assert deck.story(story.id).text == story.text


def test_format_text_may_cross_a_paragraph_boundary(deck):
    story = next(s for s in deck[0].shapes if s.name == "Title").stories[0]
    assert len(story.paragraphs) == 2

    edit = deck.format_text(story.id, 10, 20, color="#00ff00")
    assert (edit.start, edit.end) == (10, 20)

    formatted = deck.story(story.id)
    assert formatted.text == story.text
    assert [
        "".join(run.text for run in para.runs if run.color == "#00ff00")
        for para in formatted.paragraphs
    ] == ["es,", "withou"]


def test_render_slide_needs_a_registered_font(deck):
    with pytest.raises(bo.RenderError, match="font"):
        deck.render_slide(0)


def test_render_slide_produces_a_display_list(deck, font_bytes, tmp_path):
    assert deck.register_font("Inter", font_bytes) == 0
    assert deck.register_font("Inter", font_bytes, bold=True) == 1

    layout = deck.render_slide(0)
    assert (layout.width, layout.height) == (1280.0, 720.0)
    assert len(layout) > 0
    assert layout.contract_version >= 1

    scene = layout.to_dict()
    assert scene["width"] == 1280.0
    assert len(scene["primitives"]) == len(layout)

    target = tmp_path / "slide.json"
    layout.write(target)
    assert target.read_text() == layout.json


def test_render_slide_accepts_a_slide_id(deck, font_bytes):
    deck.register_font("Inter", font_bytes)
    assert len(deck.render_slide(deck.slide_ids[0])) == len(deck.render_slide(0))


def test_save_round_trips_an_unedited_deck(deck):
    saved = deck.save()
    assert saved[:4] == PPTX_MAGIC

    reopened = bo.Presentation.open(saved)
    assert reopened.slide_count == 3
    assert reopened[0].text == deck[0].text


def test_save_path(deck, tmp_path):
    target = tmp_path / "out.pptx"
    deck.save_path(target)
    assert bo.Presentation.open_path(target).slide_count == 3


def test_a_moved_shape_survives_save_and_reopen(deck):
    shape_id = deck[0].shapes[0].id
    deck.move_shape(0, shape_id, 111, 222)

    reopened = bo.Presentation.open(deck.save())
    moved = reopened[0].shapes[0]
    assert (moved.x, moved.y) == (111, 222)


def test_an_inserted_slide_survives_save_and_reopen(deck, tmp_path):
    deck.insert_slide(1)

    target = tmp_path / "out.pptx"
    deck.save_path(target)
    assert bo.Presentation.open_path(target).slide_count == 4


def test_an_added_text_box_survives_save_and_reopen(deck):
    deck.add_text_box(0, x=0, y=0, width=10, height=10, text="written back")

    reopened = bo.Presentation.open(deck.save())
    assert "written back" in reopened[0].text


def test_is_edited_reports_accepted_edits(deck):
    assert deck.is_edited is False
    assert deck.save()[:4] == PPTX_MAGIC

    deck.move_shape(0, deck[0].shapes[0].id, 5, 5)
    assert deck.is_edited is True
    assert deck.save()[:4] == PPTX_MAGIC
    with pytest.raises(AttributeError):
        deck.is_edited = False


def _title_story_id(deck):
    return next(s for s in deck[0].shapes if s.name == "Title").stories[0].id


REFUSED_EDITS = [
    ("unknown slide", KeyError, lambda deck: deck.delete_slide("nope")),
    ("unknown shape", KeyError, lambda deck: deck.move_shape(0, "nope", 1, 1)),
    ("unknown story", KeyError, lambda deck: deck.insert_text("nope", 0, "x")),
    (
        "bad geometry",
        ValueError,
        lambda deck: deck.add_shape(0, "notashape", x=0, y=0, width=10, height=10),
    ),
    (
        "bad size",
        ValueError,
        lambda deck: deck.add_shape(0, "rect", x=0, y=0, width=0, height=10),
    ),
    (
        "bad colour",
        ValueError,
        lambda deck: deck.set_shape_fill(0, deck[0].shapes[0].id, "red"),
    ),
    ("insert out of range", bo.RangeError, lambda deck: deck.insert_slide(99)),
    ("move out of range", bo.RangeError, lambda deck: deck.move_slide(0, 99)),
    (
        "text across paragraphs",
        bo.RangeError,
        lambda deck: deck.delete_text(_title_story_id(deck), 10, 20),
    ),
]


@pytest.mark.parametrize(
    ("expected", "edit"),
    [case[1:] for case in REFUSED_EDITS],
    ids=[case[0] for case in REFUSED_EDITS],
)
def test_an_edit_the_engine_refused_leaves_the_deck_unedited(
    deck, tmp_path, expected, edit
):
    """Only an edit the engine accepted marks the deck edited."""
    with pytest.raises(expected):
        edit(deck)

    assert deck.is_edited is False
    assert deck.save()[:4] == PPTX_MAGIC
    deck.save_path(tmp_path / "out.pptx")


def test_a_peer_update_marks_the_replica_edited_and_saves_the_edit(sample_bytes):
    left = bo.Presentation.open_collaborative(sample_bytes, client_id=101)
    right = bo.Presentation.open_collaborative(sample_bytes, client_id=202)
    left.move_shape(0, left[0].shapes[0].id, 3, 4)

    right.apply_update(left.diff(right.state_vector()))
    assert right.is_edited is True
    reopened = bo.Presentation.open(right.save())
    moved = reopened[0].shapes[0]
    assert (moved.x, moved.y) == (3, 4)


GIL_PROBE = """
import os, sys, threading, time

import betteroffice_pptx as bo

mode, fifo, source = sys.argv[1:4]
data = open(source, "rb").read()


def feed():
    time.sleep(0.5)
    handle = os.open(fifo, os.O_WRONLY)
    written = 0
    while written < len(data):
        written += os.write(handle, data[written:])
    os.close(handle)


def drain():
    time.sleep(0.5)
    handle = os.open(fifo, os.O_RDONLY)
    while os.read(handle, 65536):
        pass
    os.close(handle)


if mode == "read":
    threading.Thread(target=feed).start()
    assert bo.Presentation.open_path(fifo).slide_count == 3
else:
    threading.Thread(target=drain).start()
    bo.Presentation.open(data).save_path(fifo)
"""


@pytest.mark.skipif(not hasattr(os, "mkfifo"), reason="needs POSIX FIFOs")
@pytest.mark.parametrize("mode", ["read", "write"])
def test_path_io_releases_the_gil(sample_path, tmp_path, mode):
    """A FIFO opens only once the peer thread runs, which needs the GIL back."""
    fifo = tmp_path / "deck.fifo"
    os.mkfifo(fifo)

    try:
        probe = subprocess.run(
            [sys.executable, "-c", GIL_PROBE, mode, str(fifo), str(sample_path)],
            capture_output=True,
            text=True,
            timeout=30,
        )
    except subprocess.TimeoutExpired:
        pytest.fail(f"{mode} deadlocked: the GIL was held across the file IO")
    assert probe.returncode == 0, probe.stderr


def test_missing_file_raises_file_not_found(tmp_path):
    missing = tmp_path / "nope.pptx"
    with pytest.raises(FileNotFoundError) as caught:
        bo.Presentation.open_path(missing)
    assert caught.value.errno is not None
    assert caught.value.filename == str(missing)


def test_unwritable_save_path_raises_oserror(deck, tmp_path):
    target = tmp_path / "missing-dir" / "out.pptx"
    with pytest.raises(OSError) as caught:
        deck.save_path(target)
    assert caught.value.filename == str(target)


def test_garbage_bytes_raise_parse_error():
    with pytest.raises(bo.ParseError):
        bo.Presentation.open(b"not a pptx")


def test_error_hierarchy_rolls_up_to_pptx_error():
    for subclass in (
        bo.ParseError,
        bo.RangeError,
        bo.RenderError,
        bo.InvalidUpdateError,
        bo.CollaborativeStateError,
        bo.NotCollaborativeError,
    ):
        assert issubclass(subclass, bo.PptxError)
    with pytest.raises(bo.PptxError):
        bo.Presentation.open(b"not a pptx")


def test_unknown_ids_raise_key_error(deck):
    with pytest.raises(KeyError):
        deck.delete_slide("nope")
    with pytest.raises(KeyError):
        deck.move_shape(0, "nope", 1, 1)
    with pytest.raises(KeyError):
        deck.story("nope")


def test_bool_is_not_a_slide_index(deck):
    for value in (True, False):
        with pytest.raises(TypeError, match="not bool"):
            deck.slide(value)


def test_out_of_range_slide_keys_raise_index_error(deck):
    for value in (-1, 99, 10**100):
        with pytest.raises(IndexError):
            deck.slide(value)


def test_bad_slide_key_type_raises_type_error(deck):
    with pytest.raises(TypeError):
        deck.slide(1.5)


def test_out_of_bounds_indices_raise_range_error(deck):
    with pytest.raises(bo.RangeError):
        deck.insert_slide(99)
    with pytest.raises(bo.RangeError):
        deck.move_slide(0, 99)

    edit = deck.add_text_box(0, x=0, y=0, width=100, height=100, text="Hi")
    story = deck[0].shapes[edit.index].stories[0]
    with pytest.raises(bo.RangeError):
        deck.insert_text(story.id, 999, "x")


@pytest.mark.parametrize(
    "call",
    [
        lambda deck: deck.add_shape(0, "notashape", x=0, y=0, width=10, height=10),
        lambda deck: deck.add_shape(0, "rect", x=0, y=0, width=0, height=10),
        lambda deck: deck.set_shape_fill(0, deck[0].shapes[0].id, "red"),
    ],
)
def test_invalid_shape_arguments_raise_value_error(deck, call):
    with pytest.raises(ValueError):
        call(deck)


def test_unknown_adjustment_guide_is_refused(deck):
    edit = deck.add_shape(0, "roundRect", x=0, y=0, width=10, height=10)
    with pytest.raises(ValueError, match="guide name"):
        deck.set_shape_adjust(0, edit.shape_id, {"bogus": 0.5})


def test_parse_limits_are_applied_and_validated(sample_bytes):
    with pytest.raises(ValueError, match="unknown parse limit"):
        bo.Presentation.open(sample_bytes, limits={"nope": 1})
    with pytest.raises(bo.ParseError, match="limit"):
        bo.Presentation.open(sample_bytes, limits={"max_shapes": 2})
    assert bo.Presentation.open(sample_bytes, limits={"max_shapes": 1000})


def test_origin_is_validated(deck):
    assert deck.origin == "local"
    deck.origin = "agent"
    assert deck.origin == "agent"
    with pytest.raises(ValueError, match="local, agent, remote, or system"):
        deck.origin = "sideways"


def test_author_round_trips(deck):
    assert deck.author == "python"
    deck.author = "ana"
    assert deck.author == "ana"


def test_declared_version_matches_runtime():
    import importlib.metadata as metadata

    assert metadata.version("betteroffice-pptx") == bo.__version__
    assert bo.__version__.count(".") == 2


def test_emu_constants_are_the_ooxml_ones():
    assert bo.EMU_PER_INCH == 914400
    assert bo.EMU_PER_CENTIMETER == 360000
    assert bo.EMU_PER_POINT == 12700


def test_replicas_converge(sample_bytes):
    left = bo.Presentation.open_collaborative(sample_bytes, client_id=101)
    right = bo.Presentation.open_collaborative(sample_bytes, client_id=202)
    assert (left.client_id, right.client_id) == (101, 202)

    shape = left[0].shapes[0].id
    left.move_shape(0, shape, 777, 888)
    merged = right.apply_update(left.diff(right.state_vector()))

    assert (merged.slides[0].shapes[0].x, merged.slides[0].shapes[0].y) == (777, 888)
    assert (right[0].shapes[0].x, right[0].shapes[0].y) == (777, 888)


def test_a_fresh_replica_catches_up_from_one_update(sample_bytes):
    source = bo.Presentation.open_collaborative(sample_bytes, client_id=101)
    source.move_shape(0, source[0].shapes[0].id, 555, 666)

    joiner = bo.Presentation.open_collaborative(sample_bytes, client_id=303)
    joiner.apply_update(source.state_as_update())
    assert (joiner[0].shapes[0].x, joiner[0].shapes[0].y) == (555, 666)


def test_collaborative_open_generates_read_only_client_ids(sample_bytes):
    left = bo.Presentation.open_collaborative(sample_bytes)
    right = bo.Presentation.open_collaborative(sample_bytes)

    assert left.client_id != right.client_id
    assert 0 < left.client_id < 2**53 - 1
    with pytest.raises(AttributeError):
        left.client_id = 99


def test_explicit_client_ids_are_validated(sample_bytes):
    for client_id in (0, 2**60):
        with pytest.raises(ValueError, match="client ID"):
            bo.Presentation.open_collaborative(sample_bytes, client_id=client_id)


@pytest.mark.parametrize("buffer_type", [bytearray, memoryview])
def test_collaboration_accepts_bytes_like_inputs(sample_bytes, buffer_type):
    assert bo.Presentation.open(buffer_type(sample_bytes)).slide_count == 3
    left = bo.Presentation.open_collaborative(buffer_type(sample_bytes), client_id=401)
    right = bo.Presentation.open_collaborative(buffer_type(sample_bytes), client_id=402)

    left.move_shape(0, left[0].shapes[0].id, 12, 34)
    update = left.diff(buffer_type(right.state_vector()))
    right.apply_update(buffer_type(update))
    assert right[0].shapes[0].x == 12


def test_oversized_update_is_rejected_before_copy(sample_bytes):
    class CopyGuard(bytearray):
        copied = False

        def __bytes__(self):
            self.copied = True
            raise AssertionError("oversized payload was copied")

    deck = bo.Presentation.open_collaborative(sample_bytes)
    update = CopyGuard(bo.MAX_COLLABORATION_BYTES + 1)

    with pytest.raises(bo.PptxError, match="exceeds the .*byte limit"):
        deck.apply_update(update)
    assert update.copied is False


def test_collaboration_errors_are_actionable(sample_bytes):
    deck = bo.Presentation.open_collaborative(sample_bytes)
    with pytest.raises(bo.InvalidUpdateError):
        deck.apply_update(b"\xff")
    with pytest.raises(bo.InvalidUpdateError):
        deck.diff(b"\xff\xff\xff")


def test_two_standalone_decks_cannot_be_synced_into_divergence(sample_bytes):
    """`open` decks share one client ID, so syncing them would never converge."""
    left = bo.Presentation.open(sample_bytes)
    right = bo.Presentation.open(sample_bytes)
    assert left.is_collaborative is False
    assert left.client_id == right.client_id

    left.move_shape(0, left[0].shapes[0].id, 100, 100)
    right.move_shape(0, right[0].shapes[0].id, 200, 200)

    for call in (
        left.state_vector,
        left.state_as_update,
        lambda: left.diff(b""),
        lambda: right.apply_update(b"\x00"),
    ):
        with pytest.raises(bo.NotCollaborativeError, match="open_collaborative"):
            call()


def test_collaborative_decks_are_the_ones_that_converge(sample_bytes):
    left = bo.Presentation.open_collaborative(sample_bytes, client_id=101)
    right = bo.Presentation.open_collaborative(sample_bytes, client_id=202)
    assert left.is_collaborative and right.is_collaborative

    left.move_shape(0, left[0].shapes[0].id, 100, 100)
    right.move_shape(0, right[0].shapes[0].id, 200, 200)
    left.apply_update(right.diff(left.state_vector()))
    right.apply_update(left.diff(right.state_vector()))

    assert left[0].shapes[0].x == right[0].shapes[0].x
    assert left.state_vector() == right.state_vector()


def test_is_collaborative_is_read_only(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    with pytest.raises(AttributeError):
        deck.is_collaborative = True


def test_undo_redo_round_trip(deck):
    shape = deck[0].shapes[0]
    deck.move_shape(0, shape.id, 999, 999)

    assert deck.can_undo and not deck.can_redo
    assert deck.undo() is True
    assert (deck[0].shapes[0].x, deck[0].shapes[0].y) == (shape.x, shape.y)
    assert deck.can_redo

    assert deck.redo() is True
    assert deck[0].shapes[0].x == 999


def test_undo_barrier_splits_coalesced_edits(deck):
    shape = deck[0].shapes[0].id
    deck.move_shape(0, shape, 1, 1)
    deck.add_undo_barrier()
    deck.move_shape(0, shape, 2, 2)

    deck.undo()
    assert deck[0].shapes[0].x == 1


def test_undo_covers_local_edits_but_not_remote_ones(sample_bytes):
    left = bo.Presentation.open_collaborative(sample_bytes, client_id=101)
    right = bo.Presentation.open_collaborative(sample_bytes, client_id=202)
    shape = left[0].shapes[0]
    other = left[0].shapes[1]

    left.move_shape(0, shape.id, 700, 700)
    right.move_shape(0, other.id, 900, 900)
    left.apply_update(right.diff(left.state_vector()))

    left.undo()
    assert left[0].shapes[0].x == shape.x
    assert left[0].shapes[1].x == 900


def test_agent_edits_stay_out_of_local_undo(deck):
    deck.origin = "agent"
    deck.move_shape(0, deck[0].shapes[0].id, 321, 321)

    assert deck[0].shapes[0].x == 321
    assert deck.can_undo is False


def _on_a_worker_thread(work):
    thread = threading.Thread(target=work)
    thread.start()
    thread.join()


def test_touching_a_presentation_off_its_thread_escapes_except_exception(deck):
    """`unsendable` panics, and PanicException is not an `Exception`."""
    caught: list = []

    def work() -> None:
        try:
            deck.slide_count
        except Exception as error:
            caught.append(("Exception", error))
        except BaseException as error:
            caught.append(("BaseException", error))

    _on_a_worker_thread(work)

    assert len(caught) == 1, "the deck was reachable from another thread"
    where, error = caught[0]
    assert where == "BaseException"
    assert type(error).__name__ == "PanicException"
    assert not isinstance(error, Exception)
    assert deck.slide_count == 3


def test_dropping_a_presentation_off_its_thread_leaks_it(sample_bytes, monkeypatch):
    """pyo3 skips the Rust destructor and writes an unraisable error instead."""
    unraisable: list = []
    monkeypatch.setattr(sys, "unraisablehook", unraisable.append)
    holder = [bo.Presentation.open(sample_bytes)]

    _on_a_worker_thread(holder.clear)

    assert [type(hook.exc_value).__name__ for hook in unraisable] == ["RuntimeError"]
    assert "unsendable" in str(unraisable[0].exc_value)


def test_collecting_a_cycle_off_the_owning_thread_leaks_it(sample_bytes, monkeypatch):
    """One deck per worker is not enough: the collector runs where it likes."""
    gc.collect()
    unraisable: list = []
    monkeypatch.setattr(sys, "unraisablehook", unraisable.append)

    gc.disable()
    try:
        cycle: dict = {"deck": bo.Presentation.open(sample_bytes)}
        cycle["self"] = cycle
        del cycle
        _on_a_worker_thread(gc.collect)
    finally:
        gc.enable()

    assert [type(hook.exc_value).__name__ for hook in unraisable] == ["RuntimeError"]
    assert "unsendable" in str(unraisable[0].exc_value)


def test_repr_is_python_shaped(deck):
    edit = deck.add_shape(0, "rect", x=0, y=0, width=10, height=10, fill="#123456")
    fill = deck.set_shape_fill(0, edit.shape_id, "#654321")
    stroke = deck.set_shape_stroke(0, edit.shape_id, color="#000000")
    cleared = deck.set_shape_fill(0, edit.shape_id, None)

    reprs = [
        repr(deck),
        repr(deck[0]),
        repr(deck[0].shapes[0]),
        repr(fill),
        repr(stroke),
        repr(cleared),
        repr(deck.insert_slide(0)),
    ]
    for text in reprs:
        assert "Some(" not in text, text
        assert "Nothing" not in text, text
    assert "before='#123456'" in repr(fill).replace('"', "'")
    assert "after=None" in repr(cleared)
    assert "from_index=None" in repr(deck.insert_slide(0))
    assert repr(deck).startswith("Presentation(slides=")
