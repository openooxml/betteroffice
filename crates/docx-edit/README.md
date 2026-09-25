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

Hosts edit through version-checked batches: `EditingDoc::version`,
`read_paragraphs` and `find_text` return projected paragraph text with the
session version it was read at, and `apply_edits` resolves every step against
that version, stages the whole batch on a private clone, and adopts it as one
transaction (one undo step by default) or returns an `EditRefusal` with the
document, history and id allocation untouched. `validate_edits` runs the same
checks without changing anything.

`structured` exports read-only structured content and Markdown in schema
version 1: `EditingDoc::export_structured` and `export_markdown` read a live
session with its version, and `export_docx_structured` reads DOCX bytes through
the same walker. Blocks and inlines carry anchors (paragraph ids, batch-offset
ranges, table and control ids, or `sourcePart` locations in retained XML), and
everything omitted or not represented is diagnosed. `read_types` holds the
anchor, story-selection and content-control types shared by the read APIs.

`content_controls` lists every content control in document order with the
export's control metadata, placement, anchor, parent, canonical text and
effective lock: `EditingDoc::list_content_controls` and `find_content_controls`
read a live session with its version, and `list_docx_content_controls` and
`list_package_content_controls` read bytes and parsed packages. A
`SetContentControlText` batch step fills a plain- or rich-text control by engine
id or unique tag, replacing inline control content or a block control's child
story and clearing its placeholder state in the same transaction; locked, bound,
nested and non-text controls and ambiguous tags are refused as data. Text
controls hold their text as content: local embed and raw writes never introduce,
replace or move an authored text `value`, retyping a valued control drops it,
and a fill drops it outside undo history. Collaboration updates integrate
whatever values they carry; saving ignores a text control's value, and
discovery reads its content and flags the value with `legacy-control-value`.

Used by [betteroffice-docx](https://crates.io/crates/betteroffice-docx).

Measure DOCX parsing, seeding, and body lowering with:

```sh
cargo run -p betteroffice-docx-edit --example open_probe -- input.docx 3
```

The first JSON line is a warm-up; the remaining lines are measured samples.
`openMs` includes parsing and seeding. `lowerBodyMs` lowers the body and its nested
stories without font measurement or pagination. File reads, state fingerprints,
and block fingerprints are outside the timers. Compare the same build profile on
an otherwise idle host, and use an external process deadline for large files.
The SHA-256 fingerprints let before/after runs check that seeding and lowering
agree across Rust toolchain versions.

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
