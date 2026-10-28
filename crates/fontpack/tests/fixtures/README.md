# Synthetic font fixtures

`make_fixtures.py` creates `Synthetic-Regular.ttf` and `Synthetic-Bold.ttf`
using fontTools FontBuilder. The five glyphs are original rectangles, triangles,
an empty space and a composite of those shapes. No licensed outlines or corpus
files are used. Timestamps are fixed.

Generate the TTF, identity-stream `.uw2` and standard `.woff2` fixtures from the
repository root, using the pinned tooling environment in
[`scripts/fonts/README.md`](../../../../scripts/fonts/README.md):

```sh
/tmp/fontpack-venv/bin/pip install zstandard==0.25.0
/tmp/fontpack-venv/bin/python crates/fontpack/tests/fixtures/make_fixtures.py
```

Commit the six font files and the three checksum-bearing `.zst` fixtures beside
this script. The Regular frame has no dictionary, the Bold `.ttf.zst` frame
uses the Regular `.uw2` as its raw dictionary, and the Bold `.nodict.zst` frame
is independent. Rust tests also compress the `.uw2` files in memory.
TypeScript woff2t tests only require the built wasm decoder.
