# betteroffice-docx-edit

The collaborative editing schema every DOCX editing slice runs on, backed by
Yrs.

The load-bearing rule: a Word *story* is one continuous `yrs::TextRef`. A story
is the body flow, one header or footer part, one table cell, one footnote, and
so on. `StoryId` is deliberately opaque, so it can carry package relationship
IDs or structural cell IDs without changing the schema.

`EditingDoc` is the entry point. Around it sit the typed operation vocabulary,
paragraph properties and snapshots, comment anchors that survive concurrent
edits, revision and author identity for tracked changes, and `frame_delta` for
encoding what changed between two frames so the host repaints the minimum.

Because a story is plain CRDT text, concurrent edits merge in the engine rather
than on a server, and the same schema serves the native and wasm editors.

Used by [betteroffice-docx](https://crates.io/crates/betteroffice-docx).

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
