"""Checks of cjk-slices.py on a small synthetic CFF face (fontTools 4.60.1)."""

import importlib.util
import os
import struct
import tempfile
import unittest
from pathlib import Path

from fontTools.cffLib import FDArrayIndex, FDSelect, FontDict
from fontTools.fontBuilder import FontBuilder
from fontTools.pens.t2CharStringPen import T2CharStringPen

SCRIPT = Path(os.environ.get("CJK_SLICES_SCRIPT", Path(__file__).with_name("cjk-slices.py")))
SPEC = importlib.util.spec_from_file_location("cjk_slices", SCRIPT)
SLICES = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(SLICES)

CODEPOINTS = [0x4E00 + index for index in range(6)]


def synthetic_face(path):
    """A CID-keyed CFF face like the Noto CJK faces, one glyph per codepoint."""
    names = [".notdef"] + [f"cid{index:05d}" for index in range(1, len(CODEPOINTS) + 1)]
    builder = FontBuilder(1000, isTTF=False)
    builder.setupGlyphOrder(names)
    builder.setupCharacterMap({cp: names[index + 1] for index, cp in enumerate(CODEPOINTS)})
    charstrings = {}
    for index, name in enumerate(names):
        pen = T2CharStringPen(600, None)
        pen.moveTo((50, 50))
        pen.lineTo((550, 50 + index * 40))
        pen.lineTo((300, 700))
        pen.closePath()
        charstrings[name] = pen.getCharString()
    builder.setupCFF("SyntheticCJK", {"FullName": "Synthetic CJK"}, charstrings, {})
    builder.setupHorizontalMetrics({name: (600, 50) for name in names})
    builder.setupHorizontalHeader(ascent=880, descent=-120)
    builder.setupNameTable({"familyName": "Synthetic CJK", "styleName": "Regular"})
    builder.setupOS2()
    builder.setupPost()
    top = builder.font["CFF "].cff.topDictIndex[0]
    fd = FontDict()
    fd.FontName = "SyntheticCJK-Regular"
    fd.Private = top.Private
    top.ROS = ("Adobe", "Identity", 0)
    top.CIDCount = len(names)
    top.FDArray = FDArrayIndex()
    top.FDArray.append(fd)
    top.FDSelect = FDSelect()
    top.FDSelect.format = 3
    top.FDSelect.gidArray = [0] * len(names)
    top.rawDict.pop("Private", None)
    top.__dict__.pop("Private", None)
    for name in names:
        top.CharStrings[name].private = fd.Private
    builder.save(path)


def entries(data):
    (count,) = struct.unpack_from(">I", data)
    at, out = 4, []
    for _ in range(count):
        gid, length = struct.unpack_from(">II", data, at)
        out.append((gid, data[at + 8:at + 8 + length]))
        at += 8 + length
    return out


def chunk(pairs):
    return struct.pack(">I", len(pairs)) + b"".join(struct.pack(">II", gid, len(o)) + o for gid, o in pairs)


class CutTest(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.path = Path(self.directory.name) / "Synthetic-Regular.otf"
        synthetic_face(self.path)
        self.order = [[]] * SLICES.CORE_SLICES + [CODEPOINTS[:2], CODEPOINTS[2:4]]

    def tearDown(self):
        self.directory.cleanup()

    def cut_capturing(self):
        captured = []
        real = SLICES.verify
        SLICES.verify = lambda *args: (captured.append(args), real(*args))
        try:
            files, _ = SLICES.cut(self.path, self.order)
        finally:
            SLICES.verify = real
        return files, list(captured[0])

    def test_a_cut_verifies_and_reassembles(self):
        files, _ = self.cut_capturing()
        self.assertEqual(sorted(files), ["c001.bin", "c002.bin", "c003.bin", "core.bin", "skeleton.bin"])

    def test_an_outline_emitted_into_the_wrong_chunk_is_refused(self):
        files, args = self.cut_capturing()
        first, second = entries(files["c001.bin"]), entries(files["c002.bin"])
        moved = dict(files, **{"c001.bin": chunk(first[1:]), "c002.bin": chunk(second + first[:1])})
        args[4] = moved
        with self.assertRaisesRegex(ValueError, "does not carry it|not in it"):
            SLICES.verify(*args)

    def test_a_wrong_outline_in_a_chunk_is_refused(self):
        files, args = self.cut_capturing()
        first = entries(files["c001.bin"])
        swapped = [(first[0][0], first[1][1]), (first[1][0], first[0][1])] + first[2:]
        args[4] = dict(files, **{"c001.bin": chunk(swapped)})
        with self.assertRaises(ValueError):
            SLICES.verify(*args)


if __name__ == "__main__":
    unittest.main()
