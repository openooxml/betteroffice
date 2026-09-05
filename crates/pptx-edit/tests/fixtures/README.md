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
`2710a419282950442f5218975aedbd395fef64b5` (schema 5), using its locked
dependencies and a separate Cargo target directory. Copy
`generate_hidden_schema_snapshots.rs` into that checkout's
`crates/pptx-edit/examples/`, then run from that checkout:

```sh
CARGO_TARGET_DIR=/absolute/path/to/main-target cargo run --locked -p betteroffice-pptx-edit --example generate_hidden_schema_snapshots -- /absolute/path/to/this/branch
```

The generator asserts that main seeds schema 5. `deck-schema-v5.snapshot.json`
is main's serialized snapshot of the unmodified demo deck, with no hidden keys.

`deck-schema-v2-hidden.update.bin` and `deck-schema-v5-hidden.update.bin`
come from the hidden fixture with client ID 4343. Before encoding each update:

1. Add a text box to `slide:1:257`, named `Persisted v2 textbox`, at
   `(100000, 100000, 2000000, 600000)` EMU, containing `persisted on v2`.
2. Insert `edited ` at offset 0 of `story:shape:4343:0:0`.
3. Remove `slide:1:257:shape:4`.
4. Move `slide:2:258` to index 0.

Both updates contain no hidden shape-map keys. Migration must recover four
flags, preserve the edits, and leave the deleted shape absent. The v2 fixture
uses main's legacy parser without connectors, defaults the slide numbering,
and omits later shape/picture style references and theme formatting before
seeding and stamping version 2. The v5 fixture uses main's normal parser.

`deck-schema-v5-theme-hidden.update.bin` is main's seed of
`style-matrix-deck.pptx` with the first shape marked hidden and numbering set to
10. Schema 6 must retain its nonempty theme format scheme and style references
while recovering the hidden flag.

The generator also refreshes main's v2 style and connector fixtures through the
same legacy parser and defaults. The moved connector keeps `(952500, 1047750)`
and the style fixture keeps `persisted-v2 Styled`. Released v1 and historical
v3/v4 fixtures remain their independent migration oracles.

The migration-order tests observe separate transactions: legacy v2 commits
3, 4, 5, then 6; current main v5 commits only 6. Hidden keys appear only in the
schema-6 transaction. Main's schema-5 package rewrite remains a separate step.
