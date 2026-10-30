#!/usr/bin/env python3
"""Generate geometric test faces, WOFF2 companions and source hashes."""

import runpy
from pathlib import Path

import zstandard as zstd
from fontTools.fontBuilder import FontBuilder
from fontTools.pens.ttGlyphPen import TTGlyphPen

DIRECTORY = Path(__file__).resolve().parent
ROOT = Path(__file__).resolve().parents[4]
EPOCH = 2082844800


def polygon(points):
    pen = TTGlyphPen(None)
    if points:
        pen.moveTo(points[0])
        for point in points[1:]:
            pen.lineTo(point)
        pen.closePath()
    return pen.glyph()


def make_face(style, inset):
    order = [".notdef", "space", "triangle", "box", "stack"]
    glyphs = {
        ".notdef": polygon([(60, 0), (60, 700), (500, 700), (500, 0)]),
        "space": polygon([]),
        "triangle": polygon([(40 + inset, 0), (300, 700), (560 - inset, 0)]),
        "box": polygon([(80 + inset, 0), (80 + inset, 500), (520 - inset, 500), (520 - inset, 0)]),
    }
    pen = TTGlyphPen(glyphs)
    pen.addComponent("triangle", (1, 0, 0, 1, 0, 0))
    pen.addComponent("box", (0.5, 0, 0, 0.5, 150, 150))
    glyphs["stack"] = pen.glyph()
    builder = FontBuilder(1000, isTTF=True)
    builder.setupGlyphOrder(order)
    builder.setupCharacterMap({32: "space", 65: "triangle", 66: "box", 67: "stack"})
    builder.setupGlyf(glyphs)
    builder.setupHorizontalMetrics({name: (600, 0) for name in order})
    builder.setupHorizontalHeader(ascent=800, descent=-200)
    builder.setupNameTable({
        "familyName": "Synthetic",
        "styleName": style,
        "uniqueFontIdentifier": f"Synthetic-{style}-1",
        "fullName": f"Synthetic {style}",
        "psName": f"Synthetic-{style}",
        "version": "Version 1.0",
    })
    builder.setupOS2(
        sTypoAscender=800, sTypoDescender=-200,
        usWinAscent=800, usWinDescent=200,
        usWeightClass=700 if style == "Bold" else 400,
        fsSelection=0x20 if style == "Bold" else 0x40,
    )
    builder.setupPost()
    builder.setupMaxp()
    builder.font["head"].created = EPOCH
    builder.font["head"].modified = EPOCH
    builder.font["head"].macStyle = 1 if style == "Bold" else 0
    builder.font.recalcTimestamp = False
    builder.save(DIRECTORY / f"Synthetic-{style}.ttf")


def main():
    tooling = runpy.run_path(str(ROOT / "scripts/fonts/woff2-transform.py"))
    make_face("Regular", 0)
    make_face("Bold", -20)
    tooling["transform_directory"](DIRECTORY)
    regular = (DIRECTORY / "Synthetic-Regular.uw2").read_bytes()
    for style, dictionary, suffix in [
        ("Regular", None, "ttf.zst"),
        ("Bold", regular, "ttf.zst"),
        ("Bold", None, "nodict.zst"),
    ]:
        payload = (DIRECTORY / f"Synthetic-{style}.uw2").read_bytes()
        raw_dict = (
            zstd.ZstdCompressionDict(dictionary, dict_type=zstd.DICT_TYPE_RAWCONTENT)
            if dictionary else None
        )
        params = zstd.ZstdCompressionParameters.from_level(
            19, source_size=len(payload), dict_size=len(dictionary or b""),
            window_log=20, write_checksum=1, write_content_size=1, write_dict_id=0,
        )
        compressor = zstd.ZstdCompressor(compression_params=params, dict_data=raw_dict)
        frame = compressor.compress(payload)
        (DIRECTORY / f"Synthetic-{style}.{suffix}").write_bytes(frame)


if __name__ == "__main__":
    main()
