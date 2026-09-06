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

Comments use a separate shared map. Saving patches existing XML and preserves
untouched comment parts byte for byte. Deleting a thread removes its known
replies; a reply added concurrently becomes a root when its parent is absent,
so both clients and saved files retain it.

The current schema is v11. Older v1–v10 updates can be opened, and attaching the
original package with `open_from_update_with_source` imports any source comments
that the older schema did not model. Source attachment also restores run baseline
formatting on surviving text while retaining edits and explicit zero overrides.
Source attachment also restores unedited gradient outlines. Imported properties
persist in subsequent updates.
Source-free loads defer the import until the package is attached. Older clients
reject new-schema updates; collaborators must upgrade together or exchange
saved PPTX files. The existing connector, slide-number, theme-formatting,
hidden-shape, custom-geometry, comment, list-style, baseline, and gradient-outline migrations run
in order. Missing starting slide numbers default to one and stay omitted from
package JSON.

Used by [betteroffice-pptx](https://crates.io/crates/betteroffice-pptx). The
`wasm` feature exposes the JavaScript surface consumed by
[@betteroffice/pptx](https://www.npmjs.com/package/@betteroffice/pptx).

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
