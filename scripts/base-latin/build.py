import argparse
import hashlib
import json
import os
from pathlib import Path
import tempfile
from urllib.request import urlopen

import fontTools
import zstandard
from fontTools import subset
from fontTools.pens.recordingPen import DecomposingRecordingPen
from fontTools.pens.transformPen import TransformPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTFont, newTable
from fontTools.varLib.instancer import instantiateVariableFont


ROOT = Path(__file__).resolve().parents[2]
FAMILY = "BetterOffice Base Latin"
STYLES = ("Regular", "Bold", "Italic", "BoldItalic")
TIMESTAMP = 2082844800
COMMIT = "9862abc066c9b83bf84c3e8ed1fc2d8d48f9bbf0"
SOURCES = {
    "NotoSansSymbols2-Regular.ttf": (
        f"https://raw.githubusercontent.com/google/fonts/{COMMIT}/ofl/notosanssymbols2/NotoSansSymbols2-Regular.ttf",
        "7d5fb73b7ca67a6798101741f5d280a3d016a56a197afcd4199dbb57b4b82a21",
    ),
    "NotoSans.ttf": (
        f"https://raw.githubusercontent.com/google/fonts/{COMMIT}/ofl/notosans/NotoSans%5Bwdth,wght%5D.ttf",
        "bfb7bb691513f12e734dc346c03a03f784912432d7e3fa8e56efcf906fe86b3d",
    ),
}
REGULAR_RANGES = (
    (0x0020, 0x007E), (0x00A0, 0x024F), (0x0370, 0x03FF),
    (0x0400, 0x04FF), (0x2000, 0x206F), (0x20A0, 0x20CF),
    (0x2190, 0x21FF), (0x2200, 0x22FF), (0x25A0, 0x25FF),
    (0x2600, 0x26FF), (0x2700, 0x27BF), (0xFB00, 0xFB06),
)
STYLE_RANGES = (
    (0x0020, 0x007E), (0x00A0, 0x017F),
    (0x2000, 0x206F), (0x20A0, 0x20CF),
)
PUA_REMAPS = {
    0xF0B7: 0x2022,
    0xF0A7: 0x25AA,
    0xF06F: 0x25E6,
    0xF0D8: 0x27A2,
    0xF0FC: 0x2713,
    0xF076: 0x2756,
    0xF0A8: 0x25AA,
}
GAPS = {
    "NotoSans.ttf": (0x20B9,),
    "NotoSansSymbols2-Regular.ttf": (0x25C6, 0x2713, 0x2756, 0x27A2),
}
VENDORED = {
    "NotoSans.ttf": "5a859a158eb4508ace7be32cb763441b1148c48c2cfc55bb70bdbdb72dbee0e2",
    "NotoSansSymbols2-Regular.ttf": "8ddeeabbf874a3e37aa1935c2988c784d16caf1206a78ed3c9351d89ab321925",
}
SOURCES_DIR = ROOT / "scripts/base-latin/sources"
TABLES = {
    "cmap", "glyf", "loca", "head", "hhea", "hmtx", "maxp",
    "OS/2", "post", "name",
}


def source_font(filename, cache):
    url, digest = SOURCES[filename]
    asset = ROOT / "packages/fonts/assets" / filename
    path = asset if asset.exists() else cache / filename
    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        with urlopen(url, timeout=60) as response:
            data = response.read()
        if hashlib.sha256(data).hexdigest() != digest:
            raise ValueError(f"Unexpected sha256 for {filename}")
        path.write_bytes(data)
    if hashlib.sha256(path.read_bytes()).hexdigest() != digest:
        raise ValueError(f"Unexpected sha256 for {filename}")
    font = TTFont(path, recalcTimestamp=False)
    if "fvar" in font:
        font = instantiateVariableFont(
            font, {axis.axisTag: axis.defaultValue for axis in font["fvar"].axes},
            inplace=True,
        )
    return font


def vendored_path(filename):
    return SOURCES_DIR / f"{Path(filename).stem.split('[')[0]}-gap.ttf"


def refresh_sources(cache):
    SOURCES_DIR.mkdir(parents=True, exist_ok=True)
    for filename, points in GAPS.items():
        font = source_font(filename, cache)
        options = subset.Options()
        options.hinting = False
        options.layout_features = []
        options.recalc_timestamp = False
        options.notdef_outline = True
        subsetter = subset.Subsetter(options=options)
        subsetter.populate(unicodes=points)
        subsetter.subset(font)
        for tag in list(font.keys()):
            if tag != "GlyphOrder" and tag not in TABLES:
                del font[tag]
        license_text = (ROOT / "crates/base-latin/LICENSE-OFL-NotoSans").read_text()
        font["name"].setName(
            license_text[license_text.index("SIL OPEN FONT LICENSE"):].strip(), 13, 3, 1, 0x409
        )
        font["head"].created = TIMESTAMP
        font["head"].modified = TIMESTAMP
        font.recalcTimestamp = False
        destination = vendored_path(filename)
        font.save(destination, reorderTables=True)
        digest = hashlib.sha256(destination.read_bytes()).hexdigest()
        print(f"{destination.name}: {digest}")


def vendored_font(filename):
    path = vendored_path(filename)
    if hashlib.sha256(path.read_bytes()).hexdigest() != VENDORED[filename]:
        raise ValueError(f"Unexpected sha256 for {path.name}")
    return TTFont(path, recalcTimestamp=False)


def merge_glyph(target, source, codepoint):
    source_name = source.getBestCmap()[codepoint]
    scale = target["head"].unitsPerEm / source["head"].unitsPerEm
    glyphs = source.getGlyphSet()
    recording = DecomposingRecordingPen(glyphs)
    glyphs[source_name].draw(recording)
    pen = TTGlyphPen(None)
    recording.replay(TransformPen(pen, (scale, 0, 0, scale, 0, 0)))
    name = f"baseLatin{codepoint:04X}"
    order = [*target.getGlyphOrder(), name]
    target["glyf"][name] = pen.glyph()
    target.setGlyphOrder(order)
    advance, bearing = source["hmtx"].metrics[source_name]
    target["hmtx"].metrics[name] = (round(advance * scale), round(bearing * scale))
    for table in target["cmap"].tables:
        if table.isUnicode():
            table.cmap[codepoint] = name


def notices(style):
    licenses = ["LICENSE-OFL"]
    if style == "Regular":
        licenses += ["LICENSE-OFL-NotoSans", "LICENSE-OFL-NotoSansSymbols2"]
    texts = [(ROOT / "crates/base-latin" / name).read_text() for name in licenses]
    copyright = " ".join(
        " ".join(line.strip() for line in text.partition("\n\n")[0].splitlines())
        for text in texts
    )
    license_text = texts[0][texts[0].index("SIL OPEN FONT LICENSE"):].strip()
    return copyright, license_text


def rename(font, style):
    copyright, license_text = notices(style)
    font["name"] = newTable("name")
    font["name"].names = []
    subfamily = "Bold Italic" if style == "BoldItalic" else style
    names = {
        1: FAMILY, 2: subfamily,
        3: f"BetterOfficeBaseLatin-{style};1.000",
        4: f"{FAMILY} {subfamily}", 5: "Version 1.000",
        6: f"BetterOfficeBaseLatin-{style}",
        0: copyright, 13: license_text,
        14: "https://openfontlicense.org/", 16: FAMILY, 17: subfamily,
    }
    for name_id, value in sorted(names.items()):
        font["name"].setName(value, name_id, 3, 1, 0x409)


def build(output, web_output):
    if fontTools.__version__ != "4.63.0":
        raise RuntimeError("Install scripts/base-latin/requirements.txt")
    if zstandard.__version__ != "0.25.0" or zstandard.ZSTD_VERSION != (1, 5, 7):
        raise RuntimeError("Install scripts/base-latin/requirements.txt")
    output.mkdir(parents=True, exist_ok=True)
    symbols = vendored_font("NotoSansSymbols2-Regular.ttf")
    sans = vendored_font("NotoSans.ttf")
    for style in STYLES:
        font = TTFont(
            ROOT / f"packages/fonts/assets/LiberationSans-{style}.ttf",
            recalcTimestamp=False,
        )
        if style == "Regular":
            for point in (0x25C6, 0x2713, 0x2756, 0x27A2):
                merge_glyph(font, symbols, point)
            merge_glyph(font, sans, 0x20B9)
        cmap = font.getBestCmap()
        remaps = dict(PUA_REMAPS)
        ranges = REGULAR_RANGES if style == "Regular" else STYLE_RANGES
        points = {point for start, end in ranges for point in range(start, end + 1)}
        points.update(point for point in remaps.values() if point in cmap)
        options = subset.Options()
        options.hinting = False
        options.layout_features = []
        options.recalc_timestamp = False
        options.recommended_glyphs = False
        options.glyph_names = False
        subsetter = subset.Subsetter(options=options)
        subsetter.populate(unicodes=points)
        subsetter.subset(font)
        cmap = font.getBestCmap()
        for table in font["cmap"].tables:
            if table.isUnicode():
                for private, point in sorted(remaps.items()):
                    if point in cmap:
                        table.cmap[private] = cmap[point]
        for tag in list(font.keys()):
            if tag != "GlyphOrder" and tag not in TABLES:
                del font[tag]
        font["post"].formatType = 3.0
        rename(font, style)
        font["head"].created = TIMESTAMP
        font["head"].modified = TIMESTAMP
        font.recalcTimestamp = False
        destination = output / f"BetterOfficeBaseLatin-{style}.ttf"
        font.save(destination, reorderTables=True)
        print(f"{destination.name}: {destination.stat().st_size} bytes")
    web_output.mkdir(parents=True, exist_ok=True)
    faces = []
    compressor = zstandard.ZstdCompressor(level=19, write_checksum=True)
    for style in STYLES:
        source = f"BetterOfficeBaseLatin-{style}.ttf"
        data = (output / source).read_bytes()
        encoded = compressor.compress(data)
        file = f"{source}.zst"
        (web_output / file).write_bytes(encoded)
        faces.append({
            "file": file,
            "source": source,
            "family": FAMILY,
            "codec": "zstd",
            "dict": None,
            "encodedLength": len(encoded),
            "byteLength": len(data),
            "sha256": hashlib.sha256(data).hexdigest(),
        })
    manifest = {"format": 1, "schema": 1, "pack": "base-latin", "version": "1", "faces": faces}
    (web_output / "base-latin.json").write_text(json.dumps(manifest, indent=2) + "\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, default=ROOT / "crates/base-latin/fonts")
    parser.add_argument("--web-output", type=Path, default=ROOT / "packages/fonts/assets")
    parser.add_argument(
        "--cache", type=Path,
        default=Path(os.environ.get("XDG_CACHE_HOME", tempfile.gettempdir())) / "betteroffice-base-latin",
    )
    parser.add_argument("--refresh-sources", action="store_true")
    args = parser.parse_args()
    if args.refresh_sources:
        refresh_sources(args.cache)
    else:
        build(args.output, args.web_output)
