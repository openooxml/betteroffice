import io
import re
import zipfile

import pytest

import betteroffice_pptx as bo

PNG_MAGIC = b"\x89PNG\r\n\x1a\n"


def png_dimensions(png: bytes) -> tuple[int, int]:
    return (
        int.from_bytes(png[16:20], "big"),
        int.from_bytes(png[20:24], "big"),
    )


@pytest.fixture
def deck(sample_bytes, font_bytes):
    presentation = bo.Presentation.open(sample_bytes)
    presentation.register_font("Arial", font_bytes)
    presentation.register_font("Arial", font_bytes, bold=True)
    return presentation


def test_render_png_needs_a_registered_font(sample_bytes):
    presentation = bo.Presentation.open(sample_bytes)
    with pytest.raises(bo.RenderError, match="font"):
        presentation.render_png(0)


def test_render_png_produces_a_slide_sized_png(deck, tmp_path):
    png = deck.render_png(0)
    assert png.bytes[:8] == PNG_MAGIC
    assert (png.width, png.height) == (1280, 720)
    assert png_dimensions(png.bytes) == (1280, 720)
    assert png.skipped_images == 0
    assert len(png) == len(png.bytes)
    assert "1280" in repr(png)

    target = tmp_path / "slide.png"
    png.write(target)
    assert target.read_bytes() == png.bytes


def test_render_png_scales(deck):
    assert png_dimensions(deck.render_png(0, scale=2.0).bytes) == (2560, 1440)


def test_render_png_accepts_a_slide_id(deck):
    by_id = deck.render_png(deck.slide_ids[1])
    by_index = deck.render_png(1)
    assert by_id.bytes == by_index.bytes


def test_render_png_renders_every_slide(deck):
    for index in range(len(deck)):
        assert deck.render_png(index).bytes[:8] == PNG_MAGIC


def test_render_png_counts_a_picture_without_a_media_reference(sample_bytes, font_bytes):
    buffer = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(sample_bytes)) as source, zipfile.ZipFile(buffer, "w") as target:
        for part in source.infolist():
            data = source.read(part.filename)
            if part.filename == "ppt/slides/slide1.xml":
                data, removed = re.subn(rb'\s+r:embed="[^"]+"', b"", data)
                assert removed == 1
            target.writestr(part, data)

    presentation = bo.Presentation.open(buffer.getvalue())
    presentation.register_font("Arial", font_bytes)
    pictures = [
        primitive
        for primitive in presentation.render_slide(0).to_dict()["primitives"]
        if primitive["kind"] == "image"
    ]
    assert len(pictures) == 1
    assert pictures[0].get("assetId") is None
    assert presentation.render_png(0).skipped_images == 1


def test_render_png_is_deterministic(deck):
    assert deck.render_png(2).bytes == deck.render_png(2).bytes


def test_render_png_rejects_a_slide_past_the_deck(deck):
    with pytest.raises(IndexError):
        deck.render_png(99)


def test_render_png_rejects_a_scale_that_is_not_positive(deck):
    with pytest.raises(bo.PptxError, match="finite and positive"):
        deck.render_png(0, scale=0.0)


@pytest.mark.parametrize("background", ["slide", "transparent", "#102030"])
def test_render_png_accepts_every_background(deck, background):
    assert deck.render_png(0, background=background).bytes[:8] == PNG_MAGIC


def test_render_png_rejects_an_unknown_background(deck):
    with pytest.raises(ValueError, match="background must be"):
        deck.render_png(0, background="chartreuse")


def test_png_is_exported(deck):
    assert isinstance(deck.render_png(0), bo.Png)
    assert "Png" in bo.__all__
