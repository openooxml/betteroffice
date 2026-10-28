#!/usr/bin/env python3
"""Generate standard and identity-stream WOFF2 companions for TTF faces."""

import argparse
import io
from importlib.metadata import version
from pathlib import Path

from fontTools.ttLib import TTFont, woff2
from fontTools.ttLib.woff2 import WOFF2FlavorData

FONTTOOLS_VERSION = "4.60.1"
BROTLI_VERSION = "1.1.0"
EPOCH = 2082844800


class IdentityCompressor:
    MODE_FONT = 0

    @staticmethod
    def compress(data, **_options):
        return data


def encode(path, identity):
    original = woff2.brotli
    try:
        if identity:
            woff2.brotli = IdentityCompressor()
        with TTFont(path, recalcTimestamp=False) as font:
            if font.sfntVersion != "\0\1\0\0" or not {"glyf", "loca"} <= set(font.keys()):
                raise ValueError(f"Expected a TTF glyf/loca face: {path.name}")
            if "CFF " in font or "CFF2" in font:
                raise ValueError(f"CFF is unsupported: {path.name}")
            font["head"].created = EPOCH
            font["head"].modified = EPOCH
            font.flavor = "woff2"
            font.flavorData = WOFF2FlavorData(transformedTables=("glyf", "loca"))
            output = io.BytesIO()
            font.save(output, reorderTables=False)
            return output.getvalue()
    finally:
        woff2.brotli = original


def transform_directory(directory):
    if version("fonttools") != FONTTOOLS_VERSION or version("brotli") != BROTLI_VERSION:
        raise RuntimeError(f"Install fonttools=={FONTTOOLS_VERSION} brotli=={BROTLI_VERSION}")
    paths = sorted(
        path for path in directory.iterdir()
        if path.is_file() and path.suffix.lower() == ".ttf"
    )
    if not paths:
        raise ValueError("Directory has no TTF faces")
    for path in paths:
        standard = encode(path, identity=False)
        transformed = encode(path, identity=True)
        path.with_suffix(".woff2").write_bytes(standard)
        path.with_suffix(".uw2").write_bytes(transformed)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", type=Path)
    args = parser.parse_args()
    transform_directory(args.directory)


if __name__ == "__main__":
    main()
