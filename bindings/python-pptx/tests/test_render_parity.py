import base64
import io
import math
import struct
import zipfile
from pathlib import Path

import pytest

import betteroffice_pptx as bo


@pytest.fixture
def deck(sample_bytes, font_bytes):
    result = bo.Presentation.open(sample_bytes)
    for bold in (False, True):
        result.register_font("Arial", font_bytes, bold=bold)
    return result


def first_story(deck):
    return next(shape.stories[0] for shape in deck[0].shapes if shape.stories)


def diff_frame(value):
    if isinstance(value, float):
        return struct.unpack("f", struct.pack("f", value))[0]
    if isinstance(value, list):
        return [diff_frame(item) for item in value]
    if isinstance(value, dict):
        if value.get("type") in ("move", "line", "quad", "cubic", "close"):
            return value
        return {key: diff_frame(item) for key, item in value.items()}
    return value


def test_arabic_in_arial_uses_naskh_in_bundled_layout():
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    pkg = "http://schemas.openxmlformats.org/package/2006/relationships"
    p = "http://schemas.openxmlformats.org/presentationml/2006/main"
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w") as deck:
        deck.writestr("[Content_Types].xml", """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
          <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
        </Types>""")
        deck.writestr("_rels/.rels", f'<Relationships xmlns="{pkg}"><Relationship Id="doc" Type="{rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>')
        deck.writestr("ppt/_rels/presentation.xml.rels", f'<Relationships xmlns="{pkg}"><Relationship Id="slide" Type="{rel}/slide" Target="slides/slide1.xml"/></Relationships>')
        deck.writestr("ppt/presentation.xml", f'<p:presentation xmlns:p="{p}" xmlns:r="{rel}"><p:sldIdLst><p:sldId id="256" r:id="slide"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>')
        deck.writestr("ppt/slides/slide1.xml", f'''<p:sld xmlns:p="{p}" xmlns:a="{a}"><p:cSld><p:spTree>
          <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
          <p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="1200000"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400">
            <a:latin typeface="Arial"/><a:cs typeface="Arial"/>
          </a:rPr><a:t>مرحبا</a:t></a:r></a:p></p:txBody></p:sp>
        </p:spTree></p:cSld></p:sld>''')
    frame = bo.Presentation.open(output.getvalue()).render_slide(0).to_dict()
    runs = [run for item in frame["primitives"] if item["kind"] == "textBox"
            for line in item["lines"] for run in line["runs"]]
    assert {run["fontFamily"] for run in runs} == {"Noto Naskh Arabic"}


def test_hit_test_shape_text_and_empty_points(deck):
    rail = next(shape for shape in deck[0].shapes if shape.name == "Cobalt rail")
    assert deck.hit_test(0, 8.0, 8.0) == {"kind": "shape", "shapeId": rail.id}
    frame = deck.render_slide(0).to_dict()
    text = next(
        item for item in frame["primitives"]
        if item["kind"] == "textBox" and item.get("lines")
    )
    line = text["lines"][0]
    caret = line["caretStops"][0]
    expected = {
        "kind": "text", "shapeId": text["shapeId"],
        "storyId": text["storyId"], "position": caret["position"],
    }
    point = (caret["x"], line["y"] + line["height"] / 2)
    assert deck.hit_test(0, *point) == expected
    assert deck.hit_test(deck.slide_ids[0], *point) == expected
    assert deck.hit_test(0, 40.0, 20.0) is None
    for point in ((-1.0, -1.0), (1281.0, 721.0), (math.nan, 1.0), (1.0, math.inf)):
        assert deck.hit_test(0, *point) is None


@pytest.mark.without_font_packs
def test_hit_test_lays_out_without_fonts_and_tracks_edits(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    rail = next(shape for shape in deck[0].shapes if shape.name == "Cobalt rail")
    assert deck.hit_test(0, 8.0, 8.0) == {"kind": "shape", "shapeId": rail.id}
    assert deck.hit_test(0, 100.0, 150.0)["kind"] == "text"
    deck.remove_shape(0, rail.id)
    assert deck.hit_test(0, 8.0, 8.0) is None
    with pytest.raises(bo.RenderError, match="font"):
        deck.render_slide(0)


@pytest.mark.without_font_packs
def test_fallback_registration_changes_missing_glyph_layout(deck, font_bytes):
    root = Path(__file__).resolve().parents[3]
    fallback = (root / "packages/fonts/assets/NotoSansArabic-Regular.ttf").read_bytes()
    story_id = first_story(deck).id
    deck.set_story_paragraphs(story_id, [{"runs": [{
        "text": "Hello مرحبا world",
        "style": {
            "fontFamily": "Unregistered parity face", "fontSizePt": 18.0,
            "bold": False, "italic": False,
        },
    }]}])
    before = deck.render_slide(0).to_dict()
    saved = deck.save()
    font_id = deck.register_fallback_font("Noto Sans Arabic", memoryview(fallback))
    after = deck.render_slide(0).to_dict()
    assert before != after
    assert any(
        run["fontId"] == font_id
        for item in after["primitives"] if item["kind"] == "textBox"
        for line in item.get("lines", []) for run in line["runs"]
    )
    reference = bo.Presentation.open(saved)
    for bold in (False, True):
        reference.register_font("Arial", font_bytes, bold=bold)
    assert reference.register_fallback_font("Noto Sans Arabic", fallback) == font_id
    assert reference.render_slide(0).to_dict() == after
    assert deck.hit_test(0, 100.0, 70.0) == reference.hit_test(0, 100.0, 70.0)
    assert deck.save() == saved


def test_proposal_geometry_diff_matches_native_preview_frame(deck):
    shape = deck[0].shapes[0]
    proposal = deck.propose("render-parity", [{
        "type": "setShapeRect", "slideId": deck.slide_ids[0], "shapeId": shape.id,
        "rect": {
            "x": shape.x + 95_250, "y": shape.y,
            "width": shape.width, "height": shape.height,
        },
    }])
    saved = deck.save()
    before = deck[0].shapes[0].x
    result = deck.layout_proposal_diff_slide(proposal.id, 0)
    assert set(result) == {"proposal", "snapshot", "textChanges", "shapeChanges", "frame"}
    assert result["proposal"]["id"] == proposal.id
    preview = deck.preview_proposal(proposal.id).snapshot
    assert result["snapshot"] == dict(preview, slides=[preview["slides"][0]])
    assert result["textChanges"] == []
    assert result["shapeChanges"] == [{"shapeId": shape.id, "kind": "moved"}]
    assert result["frame"] == diff_frame(deck.render_proposal(proposal.id, 0).to_dict())
    assert deck.layout_proposal_diff_slide(proposal.id, deck.slide_ids[0]) == result
    assert deck[0].shapes[0].x == before
    assert deck.save() == saved


def test_proposal_text_diff_matches_native_colored_runs(deck, sample_bytes, font_bytes):
    story_id = first_story(deck).id
    style = {"fontFamily": "Arial", "fontSizePt": 18.0, "bold": False, "italic": False}
    deck.set_story_paragraphs(story_id, [{"runs": [{"text": "old", "style": style}]}])
    saved = deck.save()
    proposal = deck.propose("render-parity", [{
        "type": "replaceText", "storyId": story_id, "start": 0, "end": 3, "text": "new",
    }])
    result = deck.layout_proposal_diff_slide(proposal.id, 0)
    assert result["textChanges"] == [
        {"storyId": story_id, "start": 0, "end": 3, "kind": "deletion"},
        {"storyId": story_id, "start": 3, "end": 6, "kind": "insertion"},
    ]
    reference = bo.Presentation.open(sample_bytes)
    for bold in (False, True):
        reference.register_font("Arial", font_bytes, bold=bold)
    reference.set_story_paragraphs(story_id, [{"runs": [
        {"text": "old", "style": dict(style, color="#b91c1c")},
        {"text": "new", "style": dict(style, color="#166534")},
    ]}])
    assert result["frame"] == diff_frame(reference.render_slide(0).to_dict())
    assert deck.story(story_id).text == "old"
    assert deck.save() == saved


def test_proposal_diff_uses_browser_layout_without_fonts(sample_bytes):
    deck = bo.Presentation.open(sample_bytes)
    proposal = deck.propose("render-parity", [{
        "type": "replaceText", "storyId": first_story(deck).id,
        "start": 0, "end": 3, "text": "new",
    }])
    result = deck.layout_proposal_diff_slide(proposal.id, 0)
    assert result["frame"]["primitives"]
    assert result["textChanges"]


@pytest.mark.parametrize("slide", [-1, 99])
def test_invalid_slide_indices(deck, slide):
    proposal = deck.propose("render-parity", [{
        "type": "replaceText", "storyId": first_story(deck).id,
        "start": 0, "end": 3, "text": "new",
    }])
    with pytest.raises(IndexError):
        deck.hit_test(slide, 0.0, 0.0)
    with pytest.raises(IndexError):
        deck.layout_proposal_diff_slide(proposal.id, slide)


def test_unknown_proposal_and_invalid_fallback_fonts(deck):
    with pytest.raises(KeyError, match="no proposal unknown"):
        deck.layout_proposal_diff_slide("unknown", 0)
    for family, data in (("", b"not a font"), ("Invalid", b"not a font")):
        with pytest.raises(bo.RenderError):
            deck.register_fallback_font(family, data)
    with pytest.raises(TypeError):
        deck.hit_test(True, 0.0, 0.0)


@pytest.mark.parametrize("convert", [bytes, bytearray, memoryview])
def test_inspection_has_the_parse_json_media_contract(sample_bytes, convert):
    source = convert(sample_bytes)
    metadata = bo.inspect_presentation(source, include_media=False)
    full = bo.inspect_presentation(source, include_media=True)
    assert bo.inspect_presentation(source) == full
    deck = bo.Presentation.open(sample_bytes)
    media = deck.media()
    assert media
    assert metadata["media"] == [
        {"partPath": part.path, "contentType": part.content_type, "byteLength": len(part.bytes)}
        for part in media
    ]
    assert full["media"] == [
        {
            "partPath": part.path, "contentType": part.content_type,
            "bytes": base64.b64encode(part.bytes).decode("ascii"),
        }
        for part in media
    ]
    assert {key: value for key, value in full.items() if key != "media"} == {
        key: value for key, value in metadata.items() if key != "media"
    }
    assert len(full["slides"]) == len(deck)
    assert full["presentation"]["widthEmu"] == deck.width_emu
    assert full["presentation"]["heightEmu"] == deck.height_emu


def test_inspection_errors_and_public_exports():
    with pytest.raises(bo.ParseError):
        bo.inspect_presentation(b"not a presentation")
    with pytest.raises(TypeError):
        bo.inspect_presentation("not bytes")
    with pytest.raises(TypeError):
        bo.inspect_presentation(b"", True)
    for name in (
        "inspect_presentation", "HitTestResult", "ProposalDiffSlide", "ProposalTextChange",
        "PresentationInspection", "ShapeHit", "TextHit", "SlideDisplayList",
    ):
        assert name in bo.__all__
