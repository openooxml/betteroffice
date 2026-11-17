# Font pack tooling

Install the release tooling in a virtual environment:

```sh
python3 -m venv /tmp/fontpack-venv
/tmp/fontpack-venv/bin/pip install fonttools==4.60.1 brotli==1.1.0
/tmp/fontpack-venv/bin/python scripts/fonts/woff2-transform.py <ttf-directory>
```

The transform writes `.uw2` (identity table stream) and standard `.woff2`
companions beside each `.ttf`, with fixed head timestamps. Both transform only
`glyf` and `loca`. It also writes `<face>.source.sha256`, which lists the SHA-256
of the exact `.ttf` bytes encoded and of both companions, in `sha256sum` format.
Use `fontpack build --codec woff2t` to compress and verify them; the builder
requires a matching sidecar for every face. Regenerate companions with
`scripts/fonts/woff2-transform.py` whenever a TTF changes.
The binary is in `betteroffice-fonts`, behind its `build` feature:

```sh
cargo run -p betteroffice-fonts --features build --bin fontpack -- build --src <directory> --out <pack-directory>
```

The default builder codec remains `zstd`; CFF/OTF faces use that codec.
This tooling runs only during development or release preparation. It never
fetches dependencies or runs at package installation.

## CJK pack

Fetch the pinned, SHA-256-verified static OTF sources, then build the pack:

```sh
scripts/fonts/fetch-cjk.sh <dir>
cargo run -p betteroffice-fonts --features build --bin fontpack -- build --src <dir> --out packages/fonts-cjk/assets --pack cjk --version <fonts-cjk version> --codec zstd
```

For offline sources, pass a directory containing the five OTF basenames as the
fetch script's second argument. Local files undergo the same SHA-256 checks.
The builder requires an absent output directory; preserve `LICENSE` and
`LICENSES` outside `assets`. These CFF faces use zstd only. Each family has one
face, so the pack has no dictionaries.

Before building, cut each face into a skeleton and outline chunks (same pinned
fontTools as above):

```sh
/tmp/fontpack-venv/bin/python scripts/fonts/cjk-slices.py <dir>
/tmp/fontpack-venv/bin/python scripts/fonts/test_cjk_slices.py
```

The second command checks the cutter on a small synthetic CID-keyed face.

It writes `<face>.slices/` and `<face>.slices.sha256` beside each OTF. The
skeleton keeps every table of the face byte for byte except `CFF `, whose
CharStrings INDEX is cut off; each chunk holds desubroutinized charstrings by
glyph id, ordered by Google's frequency-ranked slicing strategies in
`nam-slices` (Apache License 2.0), with slices 1 to 20 merged into a core.
The script checks that every chunk holds the GSUB closure of its codepoints and
that the reassembled face keeps every other table and charstring; `fontpack`
then compresses the files to `assets/<face>/`, checks that its own reassembly
matches, and adds a `slices` block to the face's manifest row.
