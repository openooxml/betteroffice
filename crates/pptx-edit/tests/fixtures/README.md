# Hidden-shape fixtures

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

The schema fixtures are generated with current `origin/main` at
`54fdaa00c8242d58db61418ac3bc3b2ad6d50cb4` (schema 6), using its locked
dependencies and a separate Cargo target directory. Copy
`generate_hidden_schema_snapshots.rs` into that checkout's
`crates/pptx-edit/examples/`, then run from that checkout:

```sh
CARGO_TARGET_DIR=/absolute/path/to/main-target cargo run --locked -p betteroffice-pptx-edit --example generate_hidden_schema_snapshots -- /absolute/path/to/this/branch
```

The generator asserts that main seeds schema 6. `deck-schema-v5.snapshot.json`
is main's serialized snapshot of the unmodified demo deck, with no hidden keys.

`deck-schema-v2-hidden.update.bin` and `deck-schema-v5-hidden.update.bin`
come from the hidden fixture with client ID 4343. Before encoding each update:

1. Add a text box to `slide:1:257`, named `Persisted v2 textbox`, at
   `(100000, 100000, 2000000, 600000)` EMU, containing `persisted on v2`.
2. Insert `edited ` at offset 0 of `story:shape:4343:0:0`.
3. Remove `slide:1:257:shape:4`.
4. Move `slide:2:258` to index 0.

Both legacy updates have main’s seeded hidden shape-map keys removed before encoding; `deck-schema-v6-hidden.update.bin` retains them. Migration must recover four
flags, preserve the edits, and leave the deleted shape absent. The v2 fixture
uses main's legacy parser without connectors, defaults the slide numbering,
and omits later shape/picture style references and theme formatting before
seeding and stamping version 2. The v5 fixture uses main's normal parser and is stamped version 5.

`deck-schema-v5-theme-hidden.update.bin` is main's seed of
`style-matrix-deck.pptx` with the first shape marked hidden and numbering set to
10. Schema 6 must retain its nonempty theme format scheme and style references
while recovering the hidden flag.

The generator also refreshes main's v2 style and connector fixtures through the
same legacy parser and defaults. The moved connector keeps `(952500, 1047750)`
and the style fixture keeps `persisted-v2 Styled`. Released v1 and historical
v3/v4 fixtures remain their independent migration oracles.

The generator also opens `custom-geometry.pptx` with client ID 285. The
legacy parser/defaults produce `deck-custom-schema-v2.update.bin`; main’s
normal writer produces `deck-custom-schema-v6.update.bin`. Both omit custom
paths, which current main does not model.

The migration-order tests observe separate transactions: legacy v2 commits
3, 4, 5, 6, 7, then 8; v5 commits 6, 7, then 8; v6 commits 7 then 8; current
main v7 commits only 8. Hidden keys first appear in the schema-6 transaction and
survive schema 7’s package rewrite and schema 8’s comment metadata. Main’s
migrations retain their original version numbers.

`deck-schema-v7-comments.update.bin` is documented in `modern-comments.md`.

`deck-schema-v8-list-style.update.bin` was generated on main `1d0f41d9` from `../../pptx-render/tests/fixtures/list-style-bullets.pptx`, using `DeckSession::open` and `encode_state_as_update_v1` with client ID 29401. It exercises schema 8 to 9 migration and source reattachment when the old model did not store list styles.

`deck-schema-v9-text-overflow.update.bin` was produced on main `2c90c17f` by opening `crates/pptx-render/tests/fixtures/text-overflow.pptx` with client ID 295 and calling `encode_state_as_update_v1()`. It proves schema v11 migration preserves the old JSON until source attachment recovers explicit overflow settings.

`run-baseline-main-v9.update.bin` and `run-baseline-edited-main-v9.update.bin`
were generated on main `2c90c17f` from
`../../pptx-render/tests/fixtures/text-baseline-script.pptx` with client ID 33101.
The latter inserts `😀 ` at offset 0 of `story:slide:0:256:shape:3:0`.
They test v9 to v10 migration, source baseline recovery, and preservation of edits.

`deck-schema-v10-text-overflow.update.bin` was generated from the same deck and client ID on main `069e4d66`. It verifies the v10 baseline migration remains intact and explicit overflow settings migrate separately in v11.
