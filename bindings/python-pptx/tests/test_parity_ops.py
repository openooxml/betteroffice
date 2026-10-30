import base64
import io
import zipfile
from xml.etree import ElementTree as ET

import pytest

import betteroffice_pptx as bo
from betteroffice_pptx.edits import PptxShapeAdjustStep, PptxSlideBackgroundStep


PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a7ioAAAAASUVORK5CYII="
)
NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def _box(deck, text="Parity", **style):
    edit = deck.add_text_box(
        0, x=100_000, y=100_000, width=2_000_000, height=600_000,
        text=text, name="Parity text", **style,
    )
    shape = next(shape for shape in deck[0].shapes if shape.id == edit.shape_id)
    return shape.stories[0]


def _read_slide(deck, index=0):
    return deck.read_content()["slides"][index]


def _saved_shape(deck, name):
    reopened = bo.Presentation.open(deck.save())
    return next(shape for shape in reopened[0].shapes if shape.name == name)


def _table_deck(sample_bytes):
    cell = (
        '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Cell</a:t>'
        '</a:r></a:p></a:txBody><a:tcPr/></a:tc>'
    )
    row = '<a:tr h="400000">' + cell * 2 + '</a:tr>'
    table = (
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="10000" name="Parity table"/>'
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
    deck = bo.Presentation.open(output.getvalue())
    shape = next(shape for shape in deck[0].shapes if shape.name == "Parity table")
    return deck, shape.id


def _layout_deck(sample_bytes):
    with zipfile.ZipFile(io.BytesIO(sample_bytes)) as source:
        parts = {part.filename: source.read(part.filename) for part in source.infolist()}
    parts["ppt/slideLayouts/parityLayout.xml"] = parts["ppt/slideLayouts/slideLayout1.xml"]
    parts["ppt/slideLayouts/_rels/parityLayout.xml.rels"] = parts["ppt/slideLayouts/_rels/slideLayout1.xml.rels"]
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
        for name, data in parts.items():
            if name == "[Content_Types].xml":
                data = data.replace(b"</Types>", (
                    '<Override PartName="/ppt/slideLayouts/parityLayout.xml" '
                    'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
                    '</Types>'
                ).encode())
            elif name == "ppt/slideMasters/_rels/slideMaster1.xml.rels":
                data = data.replace(b"</Relationships>", (
                    '<Relationship Id="rIdParityLayout" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
                    'Target="../slideLayouts/parityLayout.xml"/></Relationships>'
                ).encode())
            elif name == "ppt/slideMasters/slideMaster1.xml":
                data = data.replace(b"</p:sldLayoutIdLst>", (
                    '<p:sldLayoutId id="2147483650" r:id="rIdParityLayout"/></p:sldLayoutIdLst>'
                ).encode())
            target.writestr(name, data)
    return bo.Presentation.open(output.getvalue())


def test_search_anchors_and_command_content(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    story = _box(deck, "😀 parityneedle PARITYNEEDLE")
    matches = deck.search_text("parityneedle")
    assert [match["text"] for match in matches] == ["parityneedle", "PARITYNEEDLE"]
    assert matches[0]["start"] == 3
    assert matches[0]["end"] == 15
    assert matches[0]["slideIndex"] == 0
    assert matches[0]["slideId"] == deck[0].id
    assert matches[0]["storyId"] == story.id
    assert matches[0]["shapeId"] in [shape.id for shape in deck[0].shapes]
    assert deck.search_text("parityneedle", case_sensitive=True) == matches[:1]
    assert deck.search_text("parityneedle", limit=1) == matches[:1]
    assert deck.search_text("parityneedle", limit=0) == []
    assert deck.search_text("") == []
    anchor = deck.anchor_caret(story.id, 3)
    assert isinstance(anchor, bo.CaretAnchor)
    assert deck.resolve_caret_anchor(anchor) == 3
    deck.insert_text(story.id, 0, "++")
    assert deck.resolve_caret_anchor(anchor) == 5
    assert bo.Presentation.open(sample_bytes).resolve_caret_anchor(anchor) is None
    read = deck.read_command_content({"slideIds": [deck[0].id]})
    assert read["ok"] and read["version"] == deck.version()
    assert read["slides"] == deck.read_content({"slideIds": [deck[0].id]})["slides"]
    assert read["sources"]
    for source in read["sources"]:
        assert source["storyId"] in {item["storyId"] for item in read["stories"]}
        for paragraph in source["paragraphs"]:
            assert set(paragraph) == {"paragraphId", "text", "fields", "editable"}
    assert deck.read_command_content()["ok"]
    with pytest.raises(bo.CollaborativeStateError, match="invalid deck state: missing slide missing"):
        deck.read_command_content({"slideIds": ["missing"]})
    assert "parityneedle" in _saved_shape(deck, "Parity text").text
    with pytest.raises(bo.RangeError, match="outside length"):
        deck.anchor_caret(story.id, 9999)


def test_duplicate_slide_and_history_receipts(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    assert deck.history_step("undo") == {
        "applied": False, "changedSlideIds": [],
        "deckSizeChanged": False, "slideOrderChanged": False,
    }
    original = deck[0]
    receipt = deck.duplicate_slide(original.id, 1)
    assert isinstance(receipt, bo.SlideEdit)
    assert receipt.from_index is None and receipt.to_index == 1
    assert receipt.slide_id != original.id
    assert deck[1].id == receipt.slide_id and deck[1].text == original.text
    assert bo.Presentation.open(deck.save())[1].text == original.text
    deck.add_undo_barrier()
    undo = deck.history_step("undo")
    assert undo["applied"] and not undo["deckSizeChanged"] and undo["slideOrderChanged"]
    assert set(undo["changedSlideIds"]).issubset(deck.slide_ids)
    assert deck.slide_count == 3
    redo = deck.history_step("redo")
    assert redo["applied"] and receipt.slide_id in redo["changedSlideIds"]
    assert deck.slide_count == 4
    assert type(deck.undo()) is bool
    assert type(deck.redo()) is bool
    with pytest.raises(ValueError, match="direction must be undo or redo"):
        deck.history_step("back")


def test_slide_notes_layout_and_background(sample_bytes):
    deck = _layout_deck(sample_bytes)
    slide_id = deck[0].id
    assert deck.set_slide_notes(0, "Parity speaker notes") is None
    assert _read_slide(deck)["notes"] == "Parity speaker notes"
    layout = next(layout for layout in deck.layouts if layout != deck[0].layout)
    assert deck.set_slide_layout(slide_id, layout, []) is None
    assert deck[0].layout == layout
    background = {"type": "solid", "color": {"rgb": "123456"}}
    receipt = deck.set_slide_background(0, background)
    assert receipt["slideId"] == slide_id and receipt["after"] == background
    assert _read_slide(deck)["background"] == background
    reopened = bo.Presentation.open(deck.save())
    assert reopened[0].layout == layout
    assert _read_slide(reopened)["notes"] == "Parity speaker notes"
    assert _read_slide(reopened)["background"] == background
    cleared = deck.set_slide_background(slide_id, None)
    assert cleared["before"] == background and cleared["after"] is None
    assert deck.set_slide_notes(slide_id, "") is None
    assert not _read_slide(bo.Presentation.open(deck.save())).get("notes", "")
    with pytest.raises(bo.CollaborativeStateError, match="unknown slide layout"):
        deck.set_slide_layout(0, "ppt/slideLayouts/missing.xml")
    with pytest.raises(ValueError, match="background fill is missing its colour"):
        deck.set_slide_background(0, {"type": "solid"})


def test_picture_and_media_bytes(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    for media in deck.media():
        assert deck.media_bytes(media.path) == media.bytes
    edit = deck.add_picture(
        deck[0].id, memoryview(PNG), content_type="image/png",
        x=100_000, y=200_000, width=400_000, height=400_000, name="Parity picture",
    )
    assert isinstance(edit, bo.ShapeEdit)
    picture = next(shape for shape in deck[0].shapes if shape.id == edit.shape_id)
    assert picture.kind == "picture"
    assert deck.media_bytes(f"pending-media:{picture.id}") == PNG
    reopened = bo.Presentation.open(deck.save())
    picture = next(shape for shape in reopened[0].shapes if shape.name == "Parity picture")
    assert reopened.media_bytes(picture.media_path) == PNG
    with pytest.raises(bo.CollaborativeStateError, match="media part was not found"):
        deck.media_bytes("missing")
    with pytest.raises(bo.CollaborativeStateError, match="a picture needs image data"):
        deck.add_picture(0, b"", content_type="image/png", x=0, y=0, width=1, height=1)


@pytest.mark.parametrize("method,from_index,to_index", [
    ("bring_shape_to_front", 0, -1),
    ("send_shape_to_back", -1, 0),
    ("bring_shape_forward", 0, 1),
    ("send_shape_backward", 1, 0),
])
def test_z_order(sample_bytes, method, from_index, to_index):
    deck = bo.Presentation.open(sample_bytes)
    before = [shape.id for shape in deck[0].shapes]
    shape_id = before[from_index]
    target = to_index % len(before)
    receipt = getattr(deck, method)(0, shape_id)
    assert receipt == {
        "slideId": deck[0].id, "shapeId": shape_id,
        "fromIndex": from_index % len(before), "toIndex": target,
    }
    expected = [item for item in before if item != shape_id]
    expected.insert(target, shape_id)
    assert [shape.id for shape in deck[0].shapes] == expected
    names = {shape.id: shape.name for shape in bo.Presentation.open(sample_bytes)[0].shapes}
    assert [shape.name for shape in bo.Presentation.open(deck.save())[0].shapes] == [names[item] for item in expected]
    with pytest.raises(KeyError, match="was not found"):
        getattr(deck, method)(0, "missing")


def test_story_paragraphs_alignment_and_delete_break(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    story = _box(deck)
    replacement = deck.set_story_paragraphs(story.id, [
        {"alignment": "ctr", "bullet": False, "runs": [{"text": "First", "style": {"bold": True}}]},
        {"runs": [{"text": "Second"}]},
    ])
    assert isinstance(replacement, bo.Story)
    assert replacement.text == "First\nSecond"
    assert replacement.paragraphs[0].alignment == "ctr"
    edit = deck.set_paragraph_alignment(story.id, 0, 5, "r")
    assert isinstance(edit, bo.TextEdit) and edit.story_id == story.id
    assert deck.story(story.id).paragraphs[0].alignment == "r"
    deck.set_paragraph_alignment(story.id, 0, 5, None)
    assert deck.story(story.id).paragraphs[0].alignment is None
    edit = deck.delete_paragraph_break(story.id, 5)
    assert isinstance(edit, bo.TextEdit) and edit.text == "\n"
    assert deck.story(story.id).text == "FirstSecond"
    assert _saved_shape(deck, "Parity text").text == "FirstSecond"
    with pytest.raises(ValueError, match="unrecognized paragraph alignment"):
        deck.set_paragraph_alignment(story.id, 0, 1, "sideways")
    with pytest.raises(ValueError, match="supply 1 to 100 paragraphs"):
        deck.set_story_paragraphs(story.id, [])
    with pytest.raises(bo.RangeError, match="paragraph boundary"):
        deck.delete_paragraph_break(story.id, 1)


def test_comment_position(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    comment = deck.add_comment(0, "Parity", author="Python", created="2026-10-29T00:00:00Z")
    receipt = deck.set_comment_position(comment.comment_id, 127_000, 457_200)
    assert isinstance(receipt, bo.CommentEdit)
    assert receipt.comment_id == comment.comment_id and receipt.slide_id == deck[0].id
    moved = next(item for item in deck.comments() if item.id == comment.comment_id)
    assert (moved.x, moved.y) == (127_000, 457_200)
    reopened = bo.Presentation.open(deck.save())
    assert any((item.x, item.y) == (127_000, 457_200) for item in reopened.comments())
    reply = deck.reply_to_comment(
        comment.comment_id, "Reply", author="Python", created="2026-10-29T00:01:00Z",
    )
    with pytest.raises(ValueError, match="replies share their root comment position"):
        deck.set_comment_position(reply.comment_id, 0, 0)


@pytest.mark.parametrize("axis", ["row", "column"])
@pytest.mark.parametrize("position", ["before", "after"])
def test_table_edits_and_checks(sample_bytes, axis, position):
    deck, shape_id = _table_deck(sample_bytes)
    insert = f"insert_table_{axis}"
    delete = f"delete_table_{axis}"
    version = deck.version()
    assert deck.check_table_edit(insert, 0, shape_id, 0, position, ["New A", "New B"]) is None
    assert deck.version() == version and not deck.is_edited
    receipt = getattr(deck, insert)(deck[0].id, shape_id, 0, position, ["New A", "New B"])
    assert receipt["slideId"] == deck[0].id and receipt["shapeId"] == shape_id
    assert receipt["rowCount"] == (3 if axis == "row" else 2)
    assert receipt["columnCount"] == (3 if axis == "column" else 2)
    assert len(receipt["newCellRefs"]) == 2
    assert "New A" in deck[0].text and "New B" in deck[0].text
    reopened = bo.Presentation.open(deck.save())
    assert "New A" in reopened[0].text
    table = next(shape for shape in _read_slide(reopened)["shapes"] if shape["name"] == "Parity table")["graphic"]
    assert len(table["rows"]) == receipt["rowCount"]
    assert len(table["grid"]) == receipt["columnCount"]
    index = 0 if position == "before" else 1
    assert deck.check_table_edit(delete, 0, shape_id, index) is None
    deleted = getattr(deck, delete)(0, shape_id, index)
    assert (deleted["rowCount"], deleted["columnCount"]) == (2, 2)
    assert deleted["newCellRefs"] == []
    assert "New A" not in bo.Presentation.open(deck.save())[0].text
    with pytest.raises(bo.PptxError, match=f"{axis} index 99 is out of range"):
        deck.check_table_edit(insert, 0, shape_id, 99, position)
    with pytest.raises(bo.PptxError, match=f"{axis} index 99 is out of range"):
        getattr(deck, delete)(0, shape_id, 99)


@pytest.mark.parametrize("width,expected", [(None, [500_000, 500_000, 1_000_000]),
                                            ("split", [500_000, 500_000, 1_000_000]),
                                            ("copy", [1_000_000, 1_000_000, 1_000_000])])
def test_table_column_width_policy(sample_bytes, width, expected):
    deck, shape_id = _table_deck(sample_bytes)
    assert deck.check_table_edit("insert_table_column", 0, shape_id, 0, "after", width=width) is None
    deck.insert_table_column(0, shape_id, 0, "after", width=width)
    shape = next(shape for shape in _read_slide(deck)["shapes"] if shape["id"] == shape_id)
    assert shape["graphic"]["grid"] == expected
    saved = _saved_shape(deck, "Parity table")
    assert saved.width == (3_000_000 if width == "copy" else 2_000_000)


def test_table_refusals_preserve_state(sample_bytes):
    deck, shape_id = _table_deck(sample_bytes)
    before = deck.read_content()
    for method in ("insert_table_row", "delete_table_row", "insert_table_column", "delete_table_column"):
        kwargs = {} if method.startswith("delete") else {"position": "before"}
        with pytest.raises(bo.PptxError, match="index 99 is out of range"):
            getattr(deck, method)(0, shape_id, 99, **kwargs)
    with pytest.raises(bo.PptxError, match="cells has 3 entries; supply at most 2"):
        deck.insert_table_row(0, shape_id, 0, "after", ["A", "B", "C"])
    with pytest.raises(ValueError, match="width must be split or copy"):
        deck.insert_table_column(0, shape_id, 0, "after", width="stretch")
    assert deck.read_content() == before and not deck.is_edited
    non_table = next(shape for shape in deck[0].shapes if shape.kind == "shape")
    with pytest.raises(bo.PptxError, match="shape must identify a table from the outline"):
        deck.check_table_edit("insert_table_row", 0, non_table.id, 0, "before")
    assert deck.read_content() == before and not deck.is_edited


def test_text_style_fields_and_legacy_kwargs(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    style = dict(bold=True, italic=False, font_size=14.0, color="#123456",
                 font_family="Arial", underline="sng", spacing_pt=1.5,
                 baseline_pct=20.0, kern_pt=10.0, caps="all")
    story = _box(deck, "a😀bc", **style)
    run = story.paragraphs[0].runs[0]
    for field, value in style.items():
        assert getattr(run, field) == value
    deck.format_text(story.id, 0, 5, bold=None, color=None)
    assert deck.story(story.id).paragraphs[0].runs[0].bold is True
    assert deck.story(story.id).paragraphs[0].runs[0].color == "#123456"
    legacy = dict(bold=False, italic=True, font_size=18.0, color="#654321",
                  font_family="Liberation Sans", underline="dbl")
    edit = deck.format_text(story.id, 0, 1, **legacy, spacing_pt=2.0, baseline_pct=-10.0)
    assert isinstance(edit, bo.TextEdit) and edit.text == "a"
    changed = deck.story(story.id).paragraphs[0].runs[0]
    for field, value in legacy.items():
        assert getattr(changed, field) == value
    assert changed.spacing_pt == 2.0 and changed.baseline_pct == -10.0
    deck.insert_text(story.id, 5, "z", **style)
    last = deck.story(story.id).paragraphs[0].runs[-1]
    for field, value in style.items():
        assert getattr(last, field) == value
    saved = _saved_shape(deck, "Parity text").stories[0].paragraphs[0].runs[-1]
    for field, value in style.items():
        assert getattr(saved, field) == value


def test_proposal_receipt_is_opt_in(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    story = _box(deck)
    proposal = deck.propose("python", [{"type": "replaceText", "storyId": story.id,
                                        "start": 0, "end": 6, "text": "Legacy"}])
    assert type(deck.accept_proposal(proposal.id)) is bool
    proposal = deck.propose("python", [{"type": "replaceText", "storyId": story.id,
                                        "start": 0, "end": 6, "text": "Receipt"}])
    receipt = deck.accept_proposal(proposal.id, receipt=True)
    assert receipt["proposalId"] == proposal.id and receipt["applied"] is True
    assert receipt["snapshot"]["widthEmu"] == deck.width_emu
    snapshot_shape = next(shape for shape in receipt["snapshot"]["slides"][0]["shapes"]
                          if shape["name"] == "Parity text")
    assert snapshot_shape["textStories"][0]["paragraphs"][0]["runs"][0]["text"] == "Receipt"
    assert _saved_shape(deck, "Parity text").text == "Receipt"
    with pytest.raises(KeyError, match="no proposal"):
        deck.accept_proposal("missing", receipt=True)


def test_png_shadow_budget_and_typed_steps(sample_bytes, font_bytes):
    deck = bo.Presentation.open(sample_bytes)
    deck.register_font("Arial", font_bytes)
    default = deck.render_png(0).bytes
    assert deck.render_png(0, max_shadow_pixels=134_217_728).bytes == default
    assert deck.render_png(0, max_shadow_pixels=0).bytes.startswith(b"\x89PNG\r\n\x1a\n")
    shape = deck.add_shape(0, "roundRect", x=0, y=0, width=400_000, height=400_000)
    background: PptxSlideBackgroundStep = {
        "op": "setSlideBackground", "target": {"slideId": deck[0].id},
        "background": {"type": "solid", "color": {"rgb": "123456"}},
    }
    adjust: PptxShapeAdjustStep = {
        "op": "setShapeAdjust", "target": {"slideId": deck[0].id, "shapeId": shape.shape_id},
        "adjustments": {"adj": 0.25},
    }
    assert deck.apply_edits({"expectVersion": deck.version(), "steps": [background, adjust]})["applied"]
    assert _read_slide(deck)["background"] == background["background"]
    assert next(item for item in deck[0].shapes if item.id == shape.shape_id).adjustments == {"adj": 0.25}
    with zipfile.ZipFile(io.BytesIO(deck.save())) as package:
        slide = ET.fromstring(package.read("ppt/slides/slide1.xml"))
    assert slide.find("p:cSld/p:bg/p:bgPr/a:solidFill/a:srgbClr", NS).attrib["val"] == "123456"


def test_shadow_budget_controls_raster_output(sample_bytes, font_bytes):
    seed = bo.Presentation.open(sample_bytes)
    seed.add_shape(
        0, "rect", x=1_000_000, y=1_000_000, width=1_000_000, height=1_000_000,
        fill="#FF0000", name="Parity shadow",
    )
    shadow = (
        '<a:effectLst><a:outerShdw blurRad="127000" dist="254000" dir="2700000">'
        '<a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr>'
        '</a:outerShdw></a:effectLst>'
    )
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(seed.save())) as source:
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
            for part in source.infolist():
                data = source.read(part.filename)
                if part.filename == "ppt/slides/slide1.xml":
                    xml = data.decode()
                    end = xml.index("</p:spPr>", xml.index('name="Parity shadow"'))
                    data = (xml[:end] + shadow + xml[end:]).encode()
                target.writestr(part, data)
    deck = bo.Presentation.open(output.getvalue())
    deck.register_font("Arial", font_bytes)
    assert deck.render_png(0).bytes != deck.render_png(0, max_shadow_pixels=0).bytes


@pytest.mark.parametrize("method,args", [
    ("duplicate_slide", ("missing", 0)),
    ("set_slide_notes", ("missing", "Notes")),
    ("set_slide_background", ("missing", None)),
])
def test_missing_slides_keep_rust_errors(sample_bytes, method, args):
    deck = bo.Presentation.open(sample_bytes)
    version = deck.version()
    with pytest.raises(KeyError, match='slide "missing" was not found'):
        getattr(deck, method)(*args)
    assert deck.version() == version and not deck.is_edited
