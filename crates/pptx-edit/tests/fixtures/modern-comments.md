Hand-built modern comments over `apps/demo/public/betteroffice-demo.pptx`.

The deck contains two authors, five comments across two slides, two replies in
document order opposite to timestamp order, slide and unknown anchors, a position
of `(12345, -6789)` EMU, active status, task dates, a title, completion percentage,
paragraph and line breaks, and emoji. Its third slide has no comments.

The author part includes nonempty `userId` and `providerId` values. Comment parts
use nonconventional names to exercise relationship-based discovery.

Schema reference: [MS-PPTX CT_Comment](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/161bc2c9-98fc-46b7-852b-ba7ee77e2e54).

`deck-schema-v2-comments.update.bin` was generated from this deck on main commit
`387f2392c44e31e459264663fafd65581c8346a6` with `DeckSession::open(bytes, 725)` and
`encode_state_as_update_v1()`. Its package metadata has no comment model. Migration
must import all five comments from the fingerprint-matched source package.

`deck-schema-v7-comments.update.bin` was generated from the same deck on main
commit `0c9b52ed` (schema 7) with that checkout’s own
`crates/pptx-edit/examples/schema_fixture.rs` in `seed` mode (client ID 9300)
and a separate Cargo target directory. It carries no comment model either.
Schema 8 must load it in a single migration transaction, import all five
comments once the source is attached, and save every ZIP part unchanged; the
same checkout’s `reject` mode confirms main refuses the migrated v8 update.
