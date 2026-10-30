# Synthetic font fixtures

`make_fixtures.py` creates `Synthetic-Regular.ttf` and `Synthetic-Bold.ttf`
using fontTools FontBuilder. The five glyphs are original rectangles, triangles,
an empty space and a composite of those shapes. No licensed outlines or corpus
files are used. Timestamps are fixed.

Generate the TTF, identity-stream `.uw2`, standard `.woff2` and source
`.source.sha256` fixtures from the repository root, using the pinned tooling
environment in
[`scripts/fonts/README.md`](../../../../scripts/fonts/README.md):

```sh
/tmp/fontpack-venv/bin/pip install zstandard==0.25.0
/tmp/fontpack-venv/bin/python crates/fontpack/tests/fixtures/make_fixtures.py
```

Commit the six font files, two `.source.sha256` sidecars and three checksum-bearing
`.zst` fixtures beside this script. Each sidecar contains the lowercase hexadecimal
SHA-256 of the exact source TTF bytes followed by a newline.
The Regular frame has no dictionary, the Bold `.ttf.zst` frame
uses the Regular `.uw2` as its raw dictionary, and the Bold `.nodict.zst` frame
is independent. Rust tests also compress the `.uw2` files in memory.
TypeScript woff2t tests only require the built wasm decoder.
`Synthetic-block-bomb.zst` is built byte by byte, not compressed: a frame that
declares 1,000 bytes whose single compressed block holds 100 matches of 131,074
bytes each. Both decoders must reject it without allocating the expansion.
