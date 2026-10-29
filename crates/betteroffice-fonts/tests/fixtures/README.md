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
/tmp/fontpack-venv/bin/python crates/betteroffice-fonts/tests/fixtures/make_fixtures.py
```

Commit the six font files, two `.source.sha256` sidecars and five checksum-bearing
`.zst` fixtures beside this script. Each sidecar lists the SHA-256 of the TTF,
`.uw2` and `.woff2` files in `sha256sum` format.
The Regular frame has no dictionary, the Bold `.ttf.zst` frame
uses the Regular `.uw2` as its raw dictionary, and the Bold `.nodict.zst` frame
is independent. The Regular `.sfnt.zst` frame contains its exact TTF without a
dictionary; the Bold `.sfnt.zst` frame contains its exact TTF with the Regular
TTF as its raw-content dictionary. Rust tests compare reconstructed woff2t fonts
with wuff's decode of the matching `.woff2` when `build` is enabled.
Rust tests also compress the `.uw2` files in memory.
TypeScript woff2t tests only require the built wasm decoder.
