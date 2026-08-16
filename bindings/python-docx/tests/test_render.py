import json
import math
from collections import UserDict
from pathlib import Path
from types import MappingProxyType
from typing import Any

import pytest

from betteroffice_docx import (
    MAX_PIXMAP_DIM,
    MAX_PIXMAP_PIXELS,
    DisplayList,
    Document,
    RenderError,
)

PNG_MAGIC = b"\x89PNG\r\n\x1a\n"


def _page(width: int, height: int) -> DisplayList:
    return DisplayList.from_json(
        json.dumps(
            {"pages": [{"pageIndex": 0, "width": width, "height": height, "primitives": []}]}
        )
    )


def test_layout_paginates_a_measured_envelope(
    minimal_bytes: bytes, layout_input: "dict[str, Any]"
) -> None:
    document = Document.open(minimal_bytes)
    layout = document.layout(layout_input)

    assert layout.pages == 1
    assert len(layout) == 1
    assert layout.to_dict()["pageSize"] == {"w": 816.0, "h": 1056.0}
    assert layout.to_dict()["pages"][0]["fragments"]

    display_list = layout.display_list
    assert len(display_list) == 1
    assert display_list.primitives > 0
    page = display_list.to_dict()["pages"][0]
    assert {primitive["kind"] for primitive in page["primitives"]} == {"text"}


def test_layout_takes_the_envelope_as_json_too(
    minimal_bytes: bytes, layout_input: "dict[str, Any]"
) -> None:
    document = Document.open(minimal_bytes)

    from_text = document.layout(json.dumps(layout_input))
    from_dict = document.layout(layout_input)
    assert from_text.json == from_dict.json
    assert from_text.display_list.json == from_dict.display_list.json


@pytest.mark.parametrize("mapping_type", [UserDict, MappingProxyType])
def test_layout_accepts_non_dict_mappings(
    minimal_bytes: bytes,
    layout_input: "dict[str, Any]",
    mapping_type: Any,
) -> None:
    document = Document.open(minimal_bytes)

    from_mapping = document.layout(mapping_type(layout_input))
    from_dict = document.layout(layout_input)
    assert from_mapping.json == from_dict.json
    assert from_mapping.display_list.json == from_dict.display_list.json


def test_layout_rejects_an_envelope_it_cannot_read(minimal_bytes: bytes) -> None:
    document = Document.open(minimal_bytes)

    with pytest.raises(ValueError):
        document.layout({"measured": "not a list"})
    with pytest.raises(ValueError):
        document.layout("{")


def test_a_display_list_survives_a_json_round_trip(
    minimal_bytes: bytes, layout_input: "dict[str, Any]"
) -> None:
    original = Document.open(minimal_bytes).layout(layout_input).display_list
    adopted = DisplayList.from_json(original.json)

    assert adopted.to_dict() == original.to_dict()
    assert (len(adopted), adopted.primitives) == (len(original), original.primitives)

    with pytest.raises(ValueError):
        DisplayList.from_json('{"pages": 3}')


def test_rendering_text_needs_the_family_registered(
    minimal_bytes: bytes, layout_input: "dict[str, Any]", font_bytes: bytes
) -> None:
    document = Document.open(minimal_bytes)
    display_list = document.layout(layout_input).display_list

    with pytest.raises(RenderError) as raised:
        document.render_png(display_list, 0)
    assert "font chain" in str(raised.value)

    assert document.register_font("Calibri", font_bytes) == 0
    png = document.render_png(display_list, 0)

    assert png.bytes[:8] == PNG_MAGIC
    assert len(png) == len(png.bytes) > 1000
    assert png.skipped_images == 0


def test_a_rendered_page_can_be_written(
    minimal_bytes: bytes,
    layout_input: "dict[str, Any]",
    font_bytes: bytes,
    tmp_path: Path,
) -> None:
    document = Document.open(minimal_bytes)
    document.register_font("Calibri", font_bytes)
    png = document.render_png(document.layout(layout_input).display_list)

    target = tmp_path / "page-0.png"
    png.write(target)
    assert target.read_bytes() == png.bytes


def test_rendering_refuses_a_missing_page_and_an_oversized_one(
    minimal_bytes: bytes,
) -> None:
    document = Document.open(minimal_bytes)

    with pytest.raises(RenderError) as missing:
        document.render_png(_page(64, 32), 4)
    assert "out of range" in str(missing.value)

    with pytest.raises(RenderError) as oversized:
        document.render_png(_page(MAX_PIXMAP_DIM + 1, 4), 0)
    assert str(MAX_PIXMAP_DIM) in str(oversized.value)


@pytest.mark.parametrize("page", [-1, 1 << 128, "first"])
def test_rendering_rejects_invalid_page_indices(
    minimal_bytes: bytes, page: Any
) -> None:
    document = Document.open(minimal_bytes)

    with pytest.raises(IndexError) as invalid:
        document.render_png(_page(64, 32), page)
    assert "page index" in str(invalid.value)


def test_rendering_refuses_a_page_over_the_area_budget(minimal_bytes: bytes) -> None:
    document = Document.open(minimal_bytes)
    width = math.isqrt(MAX_PIXMAP_PIXELS)
    height = MAX_PIXMAP_PIXELS // width + 1
    assert max(width, height) <= MAX_PIXMAP_DIM

    with pytest.raises(RenderError) as oversized:
        document.render_png(_page(width, height), 0)
    assert f"{MAX_PIXMAP_PIXELS}-pixel allocation cap" in str(oversized.value)


def test_concurrent_renders_are_consistent(minimal_bytes: bytes) -> None:
    from concurrent.futures import ThreadPoolExecutor

    document = Document.open(minimal_bytes)
    display_list = _page(64, 32)
    with ThreadPoolExecutor(max_workers=4) as pool:
        results = list(pool.map(lambda _: document.render_png(display_list), range(8)))

    reference = results[0]
    assert all(png.bytes == reference.bytes for png in results)
    assert all(png.skipped_images == reference.skipped_images for png in results)


def test_a_registered_image_resolves_per_part(minimal_bytes: bytes) -> None:
    document = Document.open(minimal_bytes)
    red = _solid_png()

    document.register_image("rId9", red)
    document.register_image("rId9", red, scope="header_footer", part="rId7")
    document.register_image("rId9", red, scope="footnotes")

    with pytest.raises(ValueError):
        document.register_image("rId9", red, scope="header_footer")
    with pytest.raises(ValueError):
        document.register_image("rId9", red, scope="elsewhere")
    with pytest.raises(ValueError):
        document.register_image("", red)


def test_an_unresolved_image_is_skipped_rather_than_failing_the_page(
    minimal_bytes: bytes,
) -> None:
    document = Document.open(minimal_bytes)
    display_list = DisplayList.from_json(
        json.dumps(
            {
                "pages": [
                    {
                        "pageIndex": 0,
                        "width": 64,
                        "height": 32,
                        "primitives": [
                            {"kind": "image", "relId": "rId9", "x": 0, "y": 0, "w": 16, "h": 16}
                        ],
                    }
                ]
            }
        )
    )

    assert document.render_png(display_list, 0).skipped_images == 1

    document.register_image("rId9", _solid_png())
    assert document.render_png(display_list, 0).skipped_images == 0


def test_registering_a_font_validates_its_bytes(minimal_bytes: bytes) -> None:
    document = Document.open(minimal_bytes)

    with pytest.raises(ValueError):
        document.register_font("   ", b"")
    with pytest.raises(ValueError):
        document.register_font("Calibri", b"not a font")


def _solid_png() -> bytes:
    """A 4x4 opaque red PNG."""
    return bytes.fromhex(
        "89504e470d0a1a0a0000000d49484452"
        "00000004000000040802000000269309"
        "29000000104944415478da63f8cfc000"
        "470cc47100ae930ff1385e8c11000000"
        "0049454e44ae426082"
    )
