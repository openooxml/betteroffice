These synthetic decks exercise connectors between shapes, inside nested groups,
and after the last modeled shape. They contain no external assets.

The v2 and v3 updates were regenerated with current `origin/main` at
`2b639b95d30ac8fd37fda885663c0a7dcb105943` and its locked dependencies.
Copy `generate_schema_snapshots.rs` into that checkout's
`crates/pptx-edit/examples/`, then run:

```sh
cargo run --locked -p betteroffice-pptx-edit --example generate_schema_snapshots -- /absolute/path/to/this/branch
```

Main writes schema v3. To recreate legacy v2 updates, the generator uses main's
`parse_pptx_without_connectors` and `DeckSession::from_package_with_source`, then
sets only the schema stamp to 2 with client ID 273. Thus their package and source
ordinals come from current main's legacy parser, rather than a restamped v4
package. The moved update moves `slide:0:256:shape:1` to `(952500, 1047750)`.

The v3 connector and slide-number updates come directly from main's
`DeckSession::open` with client ID 277. The v3 legacy-connector update comes from
main opening the generated v2 update with client ID 278. It is the expected
intermediate package for the composed v2 → v3 → v4 migration. Main does not store
`firstSlideNum`, so its numbered snapshot defaults to one when opened without
source bytes; new v4 snapshots preserve the parsed offset.

The release-produced v1 fixture remains unchanged.

In the basic deck, connector 3 joins shapes 2 and 4. Its transform is
`(190500, 285750, 381000, 571500)` EMU, and its line is red at 25400 EMU.

The v2 updates were regenerated again with schema-5 main at
`2710a419282950442f5218975aedbd395fef64b5` using
`generate_hidden_schema_snapshots.rs`; their bytes remain unchanged. That
generator preserves legacy connector ordinals and omits later schema fields.
The v3 fixtures retain the provenance above.
