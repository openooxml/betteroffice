# Deck schema fixtures

Deck schema 2.1 replaces the never-released 3–21 chain: `migrate_doc` carries a
released 1.0 or 2.0 snapshot forward in a single transaction. Only snapshots at
those released versions are kept here; the intermediate `v3`–`v21` seeds and the
generators that produced them were removed with that chain and remain in git
history.

`deck-schema-v1.update.bin` was produced by release `4bdccdd` and is documented
in `../schema_migration.rs`. `deck-schema-v2*.update.bin` and
`deck-custom-schema-v2.update.bin` are legacy v2 snapshots reproduced by the two
generators below.

## Hidden shapes

`hidden-shapes.pptx` is the repository's demo deck with `hidden="1"` added to
these `p:cNvPr` elements; every other ZIP part payload is unchanged:

| Slide | Snapshot shape ID | Name |
| --- | --- | --- |
| 1 | `slide:0:256:shape:0` | Cobalt rail |
| 1 | `slide:0:256:shape:8` | BetterOffice editor preview |
| 1 | `slide:0:256:shape:8.13` | PPTX tab |
| 2 | `slide:1:257:shape:4` | Format connector |
| 2 | `slide:1:257:shape:16` | Panel divider three |

Slide 1 loses 21 primitives: the cobalt rail, 14 child shapes, and six child
text boxes. Thirteen children have no hidden flag of their own. Slide 2 loses
the two marked shapes; slide 3 is unchanged.

`deck-schema-v2-hidden.update.bin` comes from the hidden fixture with client ID
4343. Before encoding the update:

1. Add a text box to `slide:1:257`, named `Persisted v2 textbox`, at
   `(100000, 100000, 2000000, 600000)` EMU, containing `persisted on v2`.
2. Insert `edited ` at offset 0 of `story:shape:4343:0:0`.
3. Remove `slide:1:257:shape:4`.
4. Move `slide:2:258` to index 0.

The update uses main's legacy parser without connectors, defaults the slide
numbering, omits shape/picture style references and theme formatting, drops
main's seeded hidden shape-map keys, and is stamped version 2. Migration must
recover four flags, preserve the edits, and leave the deleted shape absent.

`deck-schema-v2.update.bin` is the same legacy seed of `style-matrix-deck.pptx`
with `persisted-v2 ` inserted into its first story. `deck-custom-schema-v2.update.bin`
is the legacy seed of `custom-geometry.pptx` with client ID 285; it omits custom
paths, which the writer does not model.

The generator runs against `origin/main` at
`54fdaa00c8242d58db61418ac3bc3b2ad6d50cb4` (schema 6), using its locked
dependencies and a separate Cargo target directory. Copy
`generate_hidden_schema_snapshots.rs` into that checkout's
`crates/pptx-edit/examples/`, then run from that checkout:

```sh
CARGO_TARGET_DIR=/absolute/path/to/main-target cargo run --locked -p betteroffice-pptx-edit --example generate_hidden_schema_snapshots -- /absolute/path/to/this/branch
```

The generator asserts that main seeds schema 6 before restamping.

## Connectors and comments

`connectors.md` documents the connector decks and
`generate_schema_snapshots.rs`; `modern-comments.md` documents
`modern-comments.pptx` and `deck-schema-v2-comments.update.bin`.

## Composite source decks

`blip-shadow.pptx`, `chart-text-overflow.pptx`, `metafile-tracking.pptx`, and
`run-spacing-shadow.pptx` were built to pair with removed intermediate
snapshots. They stay as ready-made decks that combine bitmap effects and shape
shadows, chart fills and explicit overflow, OLE previews and character spacing,
and shadows and character spacing respectively.
