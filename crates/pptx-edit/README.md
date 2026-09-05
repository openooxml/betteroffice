# betteroffice-pptx-edit

The collaborative PPTX deck model, backed by Yrs.

`DeckSession` owns the shared document. Slide order, slides, shapes, and text
stories are separate shared types, so reordering slides does not conflict with
editing text on one of them. Text stories reuse the same one-story-is-one-CRDT-
text rule as the DOCX editing core.

`DeckUndoManager` tracks local user-origin transactions only; remote updates
stay out of local history.

State vectors, diffs, and updates are standard Yrs v1, so any transport that
speaks Yjs sync-v1 works.

New snapshots use schema v4. Opening a v1 or v2 update first applies the v3
connector migration, preserving legacy source ordinals, then the v4 slide-number
migration. A v3 update applies only the v4 migration. Missing starting numbers
default to one and stay omitted from package JSON; older readers reject v4.

Used by [betteroffice-pptx](https://crates.io/crates/betteroffice-pptx). The
`wasm` feature exposes the JavaScript surface consumed by
[@betteroffice/pptx](https://www.npmjs.com/package/@betteroffice/pptx).

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
