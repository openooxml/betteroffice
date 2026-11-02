#!/usr/bin/env python3
"""Cut CJK OTF faces into a skeleton and frequency-ordered outline chunks.

The skeleton is the face with every table byte-identical except `CFF `, whose
charstrings are all `endchar` and whose CharStrings INDEX is the last structure
of the last table, cut off: shaping, metrics and line breaks from a skeleton
equal the whole face's. A chunk holds desubroutinized charstrings by glyph id;
appending a CharStrings INDEX built from any set of chunks to the skeleton
prefix reassembles a valid face (`reassemble`, mirrored by the Rust reader in
`crates/betteroffice-fonts/src/slices.rs`).

Each slice of Google's frequency-ordered slicing strategy (vendored in
`scripts/fonts/nam-slices`) becomes a chunk; slices 1 to 20 form the core. A
chunk holds the relaxed GSUB closure of its codepoints, so every glyph shaping
can produce from them is in the chunks they map to.

Writes `<stem>.slices/` (skeleton.bin, core.bin, cNNN.bin) and
`<stem>.slices.sha256` beside each face for `fontpack build`.
"""

import argparse
import hashlib
import io
import re
import struct
from importlib.metadata import version
from pathlib import Path

from fontTools.ttLib import TTFont

FONTTOOLS_VERSION = "4.60.1"
FORMAT = 1
MAGIC = b"BOSK"
CORE_SLICES = 20
UNSLICED_CHUNK = 223
PINNED = (0x20, 0x2010, 0x25CC)
NO_CHUNK = 0xFFFF
ENDCHAR = b"\x0e"
REGIONS = {
    "NotoSansJP": "japanese",
    "NotoSansKR": "korean",
    "NotoSansSC": "simplified-chinese",
    "NotoSansTC": "traditional-chinese",
    "NotoSerifSC": "simplified-chinese",
}


def slices(path):
    """Slices from highest to lowest priority (the files list them the other way)."""
    text = path.read_text(encoding="utf-8")
    blocks = text.split("subsets {")[1:]
    order = [[int(value) for value in re.findall(r"codepoints: (\d+)", block)] for block in blocks]
    return list(reversed(order))


def gsub_edges(font):
    edges = {}

    def add(source, target):
        edges.setdefault(source, set()).add(target)

    if "GSUB" not in font:
        return edges
    for lookup in font["GSUB"].table.LookupList.Lookup:
        for subtable in lookup.SubTable:
            kind = lookup.LookupType
            if kind == 7:
                kind = subtable.ExtSubTable.LookupType
                subtable = subtable.ExtSubTable
            if kind == 1:
                for source, target in subtable.mapping.items():
                    add(source, target)
            elif kind == 2:
                for source, targets in subtable.mapping.items():
                    for target in targets:
                        add(source, target)
            elif kind == 3:
                for source, targets in subtable.alternates.items():
                    for target in targets:
                        add(source, target)
            elif kind == 4:
                for first, ligatures in subtable.ligatures.items():
                    for ligature in ligatures:
                        for component in [first, *ligature.Component]:
                            add(component, ligature.LigGlyph)
            elif kind == 8:
                for source, target in zip(subtable.Coverage.glyphs, subtable.Substitute):
                    add(source, target)
    return edges


def codepoint_glyphs(font):
    """Glyph names per codepoint: the cmap glyph and every variation-sequence glyph of it."""
    best = font.getBestCmap()
    glyphs = {cp: {name} for cp, name in best.items()}
    for table in font["cmap"].tables:
        if table.format != 14:
            continue
        for selector, pairs in table.uvsDict.items():
            for base, name in pairs:
                if base not in glyphs:
                    continue
                glyphs[base].add(name or best[base])
    return glyphs


def closure(start, edges):
    seen = set(start)
    stack = list(start)
    while stack:
        for target in edges.get(stack.pop(), ()):
            if target not in seen:
                seen.add(target)
                stack.append(target)
    return seen


def chunk_codepoints(face_codepoints, order):
    """Chunks of codepoints: core, then each later slice, then unsliced codepoints in order."""
    assigned = set()
    chunks = [[]]
    for index, members in enumerate(order):
        picked = [cp for cp in members if cp in face_codepoints and cp not in assigned]
        assigned.update(picked)
        if index < CORE_SLICES:
            chunks[0].extend(picked)
        elif picked:
            chunks.append(picked)
    rest = sorted(face_codepoints - assigned)
    chunks.extend(rest[index:index + UNSLICED_CHUNK] for index in range(0, len(rest), UNSLICED_CHUNK))
    pinned = [cp for cp in PINNED if cp in face_codepoints]
    chunks = [chunks[0] + [cp for cp in pinned if cp not in chunks[0]]] + [
        [cp for cp in codepoints if cp not in pinned] for codepoints in chunks[1:]]
    return [chunks[0]] + [codepoints for codepoints in chunks[1:] if codepoints]


def read_dict(data, start, end):
    """(operator, [(value, position, size)]) entries of a CFF DICT."""
    entries = []
    operands = []
    index = start
    while index < end:
        byte = data[index]
        if byte <= 21:
            if byte == 12:
                operator = (12, data[index + 1])
                index += 2
            else:
                operator = byte
                index += 1
            entries.append((operator, operands))
            operands = []
        elif byte == 28:
            operands.append((struct.unpack(">h", data[index + 1:index + 3])[0], index, 3))
            index += 3
        elif byte == 29:
            operands.append((struct.unpack(">i", data[index + 1:index + 5])[0], index, 5))
            index += 5
        elif byte == 30:
            position = index
            index += 1
            while True:
                nibbles = data[index]
                index += 1
                if (nibbles >> 4) == 0xF or (nibbles & 0xF) == 0xF:
                    break
            operands.append((None, position, index - position))
        elif 32 <= byte <= 246:
            operands.append((byte - 139, index, 1))
            index += 1
        elif 247 <= byte <= 250:
            operands.append(((byte - 247) * 256 + data[index + 1] + 108, index, 2))
            index += 2
        elif 251 <= byte <= 254:
            operands.append((-(byte - 251) * 256 - data[index + 1] - 108, index, 2))
            index += 2
        else:
            raise ValueError(f"reserved DICT byte {byte}")
    return entries


def read_index(data, offset):
    """(item ranges, end) of a CFF INDEX at `offset`."""
    count = struct.unpack(">H", data[offset:offset + 2])[0]
    if count == 0:
        return [], offset + 2
    size = data[offset + 2]
    base = offset + 3
    offsets = [
        int.from_bytes(data[base + index * size:base + (index + 1) * size], "big")
        for index in range(count + 1)
    ]
    origin = base + (count + 1) * size - 1
    return [(origin + offsets[i], origin + offsets[i + 1]) for i in range(count)], origin + offsets[-1]


OFFSET_OPERATORS = (15, 17, (12, 36), (12, 37))


def operator_bytes(operator):
    return bytes(operator) if isinstance(operator, tuple) else bytes([operator])


def int5(value):
    return b"\x1d" + struct.pack(">i", value)


def dict_with_offsets(data, start, end, offsets):
    """A DICT copied from `data[start:end]` with `offsets` ({operator: value}) as 5-byte operands."""
    out = b""
    for operator, operands in read_dict(data, start, end):
        if operator in offsets and operator == 18:
            size = operands[0]
            out += data[size[1]:size[1] + size[2]] + int5(offsets[operator])
        elif operator in offsets:
            out += int5(offsets[operator])
        else:
            if operator == 19:
                raise ValueError("local subroutines survived desubroutinization")
            out += b"".join(data[position:position + size] for _, position, size in operands)
        out += operator_bytes(operator)
    return out


def build_index(items):
    if not items:
        return b"\0\0"
    offsets = [1]
    for item in items:
        offsets.append(offsets[-1] + len(item))
    size = max(1, (offsets[-1].bit_length() + 7) // 8)
    return (struct.pack(">HB", len(items), size)
            + b"".join(value.to_bytes(size, "big") for value in offsets) + b"".join(items))


def cff_prefix(cff):
    """The CFF table laid out with its CharStrings INDEX last, cut off; the INDEX start."""
    data = bytes(cff)
    header = data[:data[2]]
    names, after_names = read_index(data, len(header))
    tops, after_tops = read_index(data, after_names)
    _, after_strings = read_index(data, after_tops)
    _, after_gsubrs = read_index(data, after_strings)
    top_start, top_end = tops[0]
    top = dict(read_dict(data, top_start, top_end))
    if (12, 30) not in top:
        raise ValueError("expected a CID-keyed CFF")
    charset = top[15][0][0]
    fd_select = top[(12, 37)][0][0]
    charstrings = top[17][0][0]
    fd_array = top[(12, 36)][0][0]
    _, charstrings_end = read_index(data, charstrings)
    fd_dicts, fd_end = read_index(data, fd_array)
    privates = [dict(read_dict(data, start, end))[18] for start, end in fd_dicts]
    if not (after_gsubrs <= charset < fd_select < charstrings < charstrings_end <= fd_array):
        raise ValueError("unexpected CFF structure order")
    charset_bytes = data[charset:fd_select]
    fd_select_bytes = data[fd_select:charstrings]
    private_bytes = [data[offset[0]:offset[0] + size[0]] for size, offset in privates]

    def layout(offsets):
        top_dict = dict_with_offsets(data, top_start, top_end, offsets["top"])
        fds = [dict_with_offsets(data, start, end, {18: value}) for (start, end), value in zip(fd_dicts, offsets["private"])]
        head = header + data[len(header):after_names] + build_index([top_dict]) + data[after_tops:after_gsubrs]
        out = head + charset_bytes + fd_select_bytes
        fd_index = build_index(fds)
        positions = {15: len(head), (12, 37): len(head) + len(charset_bytes), (12, 36): len(out)}
        out += fd_index
        private_positions = []
        for block in private_bytes:
            private_positions.append(len(out))
            out += block
        positions[17] = len(out)
        return out, positions, private_positions

    zero = {"top": {operator: 0 for operator in OFFSET_OPERATORS}, "private": [0] * len(privates)}
    _, positions, private_positions = layout(zero)
    prefix, check, check_private = layout({"top": positions, "private": private_positions})
    if check != positions or check_private != private_positions:
        raise ValueError("CFF layout did not converge")
    return prefix, len(prefix)


def pad4(data):
    return data + b"\0" * (-len(data) % 4)


def checksum(data):
    data = pad4(data)
    return sum(struct.unpack(f">{len(data) // 4}I", data)) & 0xFFFFFFFF


def skeleton_sfnt(source, cff):
    """sfnt bytes up to the CharStrings INDEX: every table verbatim, `cff` last."""
    prefix, _ = cff_prefix(cff)
    count = struct.unpack(">H", source[4:6])[0]
    records = {}
    for index in range(count):
        tag, _, offset, length = struct.unpack(">4sIII", source[12 + index * 16:28 + index * 16])
        records[tag] = source[offset:offset + length]
    records[b"head"] = records[b"head"][:8] + b"\0\0\0\0" + records[b"head"][12:]
    records[b"CFF "] = prefix
    tags = sorted(records)
    directory_size = 12 + 16 * len(tags)
    layout = {}
    body = b""
    for tag in [tag for tag in tags if tag != b"CFF "] + [b"CFF "]:
        layout[tag] = (directory_size + len(body), len(records[tag]))
        body += records[tag] if tag == b"CFF " else pad4(records[tag])
    header = source[:12]
    directory = b""
    for tag in tags:
        offset, length = layout[tag]
        value = 0 if tag == b"CFF " else checksum(records[tag])
        directory += struct.pack(">4sIII", tag, value, offset, length)
    cff_record = 12 + 16 * tags.index(b"CFF ")
    return header + directory + body, cff_record, layout[b"CFF "][0], layout[b"head"][0]


def charstrings_index(num_glyphs, outlines):
    data = b"".join(outlines.get(gid, ENDCHAR) for gid in range(num_glyphs))
    offsets = [1]
    for gid in range(num_glyphs):
        offsets.append(offsets[-1] + len(outlines.get(gid, ENDCHAR)))
    return struct.pack(">HB", num_glyphs, 4) + b"".join(struct.pack(">I", value) for value in offsets) + data


def reassemble(skeleton, chunks):
    """The face from a skeleton container and chunk files; mirrors the Rust reader."""
    (magic, fmt, _, num_glyphs, cff_record, cff_offset, head_offset, prefix_length, runs) = struct.unpack(
        ">4sHHIIIIII", skeleton[:32])
    if magic != MAGIC or fmt != FORMAT:
        raise ValueError("not a slice skeleton")
    prefix = skeleton[len(skeleton) - prefix_length:]
    outlines = {}
    for chunk in chunks:
        count = struct.unpack(">I", chunk[:4])[0]
        position = 4
        for _ in range(count):
            gid, length = struct.unpack(">II", chunk[position:position + 8])
            outlines[gid] = chunk[position + 8:position + 8 + length]
            position += 8 + length
    index = charstrings_index(num_glyphs, outlines)
    font = bytearray(prefix + index)
    cff_length = len(font) - cff_offset
    font += b"\0" * (-len(font) % 4)
    struct.pack_into(">II", font, cff_record + 4, checksum(font[cff_offset:cff_offset + cff_length]), cff_offset)
    struct.pack_into(">I", font, cff_record + 12, cff_length)
    struct.pack_into(">I", font, head_offset + 8, 0)
    struct.pack_into(">I", font, head_offset + 8, (0xB1B0AFBA - checksum(font)) & 0xFFFFFFFF)
    return bytes(font)


def chunk_file(outlines, gids):
    parts = [struct.pack(">I", len(gids))]
    for gid in sorted(gids):
        parts.append(struct.pack(">II", gid, len(outlines[gid])) + outlines[gid])
    return b"".join(parts)


def cut(path, order):
    source = path.read_bytes()
    font = TTFont(io.BytesIO(source), lazy=False)
    order_glyphs = font.getGlyphOrder()
    gid_of = {name: gid for gid, name in enumerate(order_glyphs)}
    num_glyphs = len(order_glyphs)

    desub = TTFont(io.BytesIO(source), lazy=False)
    cff_set = desub["CFF "].cff
    cff_set.desubroutinize()
    charstrings = cff_set.topDictIndex[0].CharStrings
    outlines = {}
    for name in order_glyphs:
        charstring = charstrings[name]
        charstring.compile()
        outlines[gid_of[name]] = bytes(charstring.bytecode)

    edges = gsub_edges(font)
    cp_glyphs = codepoint_glyphs(font)
    chunks = chunk_codepoints(set(cp_glyphs), order)
    glyph_sets = []
    for index, codepoints in enumerate(chunks):
        names = set().union(*(cp_glyphs[cp] for cp in codepoints)) if codepoints else set()
        gids = {gid_of[name] for name in closure(names, edges)}
        if index == 0:
            gids.add(0)
        glyph_sets.append(gids)
    gid_map = [NO_CHUNK] * num_glyphs
    for index, gids in enumerate(glyph_sets):
        for gid in gids:
            if gid_map[gid] == NO_CHUNK:
                gid_map[gid] = index

    skeleton_font = TTFont(io.BytesIO(source), lazy=False)
    skeleton_cff = skeleton_font["CFF "].cff
    skeleton_cff.desubroutinize()
    for name in order_glyphs:
        skeleton_cff.topDictIndex[0].CharStrings[name].program = ["endchar"]
    buffer = io.BytesIO()
    skeleton_cff.compile(buffer, skeleton_font)
    prefix, cff_record, cff_offset, head_offset = skeleton_sfnt(source, buffer.getvalue())

    runs = []
    chunk_of_cp = {cp: index for index, codepoints in enumerate(chunks) for cp in codepoints}
    for cp in sorted(chunk_of_cp):
        chunk = chunk_of_cp[cp]
        if runs and runs[-1][0] + runs[-1][1] == cp and runs[-1][2] == chunk:
            runs[-1][1] += 1
        else:
            runs.append([cp, 1, chunk])
    skeleton = (
        struct.pack(">4sHHIIIIII", MAGIC, FORMAT, 0, num_glyphs, cff_record, cff_offset, head_offset,
                    len(prefix), len(runs))
        + struct.pack(f">{num_glyphs}H", *gid_map)
        + b"".join(struct.pack(">IIHH", start, count, chunk, 0) for start, count, chunk in runs)
        + prefix
    )
    files = {"skeleton.bin": skeleton, "core.bin": chunk_file(outlines, glyph_sets[0])}
    for index in range(1, len(glyph_sets)):
        files[f"c{index:03d}.bin"] = chunk_file(outlines, glyph_sets[index])
    verify(source, font, outlines, glyph_sets, files, edges, cp_glyphs, chunks, gid_of)
    return files, reassemble(skeleton, [data for name, data in files.items() if name != "skeleton.bin"])


def verify(source, font, outlines, glyph_sets, files, edges, cp_glyphs, chunks, gid_of):
    skeleton = files["skeleton.bin"]
    pieces = [data for name, data in files.items() if name != "skeleton.bin"]
    full = TTFont(io.BytesIO(reassemble(skeleton, pieces)), lazy=False)
    for tag in font.reader.keys():
        if tag in ("CFF ", "head"):
            continue
        if full.reader[tag] != font.reader[tag]:
            raise ValueError(f"table {tag} differs after reassembly")
    if full.reader["head"][:8] + full.reader["head"][12:] != font.reader["head"][:8] + font.reader["head"][12:]:
        raise ValueError("head differs after reassembly beyond checkSumAdjustment")
    included = set().union(*glyph_sets)
    reassembled = full["CFF "].cff.topDictIndex[0].CharStrings
    for name, gid in gid_of.items():
        charstring = reassembled[name]
        charstring.compile()
        expected = outlines[gid] if gid in included else ENDCHAR
        if bytes(charstring.bytecode) != expected:
            raise ValueError(f"charstring of {name} differs after reassembly")
    for partial in ([], pieces[:1]):
        TTFont(io.BytesIO(reassemble(skeleton, partial)), lazy=False)["CFF "].cff.topDictIndex[0].CharStrings
    chunk_of_glyph = {}
    for index, gids in enumerate(glyph_sets):
        for gid in gids:
            chunk_of_glyph.setdefault(gid, set()).add(index)
    for index, codepoints in enumerate(chunks):
        names = set().union(*(cp_glyphs[cp] for cp in codepoints)) if codepoints else set()
        for name in closure(names, edges):
            if index not in chunk_of_glyph.get(gid_of[name], ()):
                raise ValueError(f"glyph {name} reachable from chunk {index} is not in it")


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("source", type=Path, help="directory of the CJK OTF faces")
    parser.add_argument("--slices", type=Path, default=Path(__file__).with_name("nam-slices"))
    args = parser.parse_args()
    if version("fonttools") != FONTTOOLS_VERSION:
        raise SystemExit(f"fonttools {FONTTOOLS_VERSION} is required, found {version('fonttools')}")
    for path in sorted(args.source.glob("*.otf")):
        family = path.stem.rsplit("-", 1)[0]
        region = REGIONS.get(family)
        if region is None:
            continue
        order = slices(args.slices / f"{region}_default.txt")
        files, full = cut(path, order)
        out = path.with_suffix(".slices")
        out.mkdir(exist_ok=True)
        for stale in out.glob("*.bin"):
            stale.unlink()
        lines = [f"{hashlib.sha256(path.read_bytes()).hexdigest()}  {path.name}"]
        for name, data in sorted(files.items()):
            (out / name).write_bytes(data)
            lines.append(f"{hashlib.sha256(data).hexdigest()}  {out.name}/{name}")
        lines.append(f"{hashlib.sha256(full).hexdigest()}  full")
        path.with_suffix(".slices.sha256").write_text("\n".join(lines) + "\n", encoding="utf-8")
        print(f"{path.name}: {len(files) - 1} chunk files, skeleton {len(files['skeleton.bin'])} bytes")


if __name__ == "__main__":
    main()
