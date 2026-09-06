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

The current schema is v10. Older v1–v9 updates can be opened. Attaching the
original package with `open_from_update_with_source` imports source comments,
list styles, and explicit text overflow settings omitted by older schemas.
Import is deterministic and runs once. Source-free loads retain their existing
model until the package is attached. Older clients reject new-schema updates;
collaborators must upgrade together or exchange saved PPTX files.

Migrations run in version order and preserve existing content. Default text
overflow settings and starting slide numbers remain omitted from package JSON.

Used by [betteroffice-pptx](https://crates.io/crates/betteroffice-pptx). The
`wasm` feature exposes the JavaScript surface consumed by
[@betteroffice/pptx](https://www.npmjs.com/package/@betteroffice/pptx).

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
