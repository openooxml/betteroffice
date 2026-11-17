import pytest

import betteroffice_vsdx as bo

PNG_MAGIC = b"\x89PNG\r\n\x1a\n"


def png_dimensions(png):
    return tuple(int.from_bytes(png[at:at + 4], "big") for at in (16, 20))


def test_png_result_and_write(foundation_path, tmp_path):
    diagram = bo.Diagram.open_path(foundation_path)
    png = diagram.render_png({"page": 0})
    assert isinstance(png, bo.Png)
    assert png.bytes[:8] == PNG_MAGIC
    assert png_dimensions(png.bytes) == (png.width, png.height)
    assert png.width > 0 and png.height > 0
    assert png.skipped_images == 0
    assert len(png) == len(png.bytes)
    path = tmp_path / "page.png"
    png.write(path)
    assert path.read_bytes() == png.bytes


def test_png_scale_and_determinism(foundation_path):
    diagram = bo.Diagram.open_path(foundation_path)
    first = diagram.render_png({"page": 0})
    second = diagram.render_png({"page": 0}, scale=2.0)
    assert (second.width, second.height) == (first.width * 2, first.height * 2)
    assert png_dimensions(second.bytes) == (second.width, second.height)
    assert first.bytes == diagram.render_png({"page": 0}).bytes


def test_png_rejects_page_out_of_range(foundation_path):
    with pytest.raises(bo.RenderError):
        bo.Diagram.open_path(foundation_path).render_png({"page": 99})


@pytest.mark.parametrize("scale", [0.0, -1.0, float("nan"), float("inf")])
def test_png_rejects_invalid_scale(foundation_path, scale):
    with pytest.raises(bo.RenderError):
        bo.Diagram.open_path(foundation_path).render_png({"page": 0}, scale=scale)


@pytest.mark.parametrize("target", [None, 0, {}, {"page": -1}, {"page": True}, {"page": "0"}, {"slide": 0}, {"page": 0, "extra": 1}])
def test_png_rejects_invalid_target(foundation_path, target):
    with pytest.raises(bo.RangeError):
        bo.Diagram.open_path(foundation_path).render_png(target)
