# PPTX engine plan

## Ownership

Rust remains authoritative for OPC access, PresentationML parsing and saving, the yrs deck, edit operations and undo, text shaping, slide layout, display-list emission, and hit testing. TypeScript only loads wasm, decodes the boundary, replays display lists, exposes public types, and supplies React chrome.

The hard fence is `crates/docx-edit/**`, `packages/docx/src/yrs/**`, `packages/docx/src/collaboration/**`, `packages/docx-react/**`, and `packages/xlsx*/**`. Those paths may be read but will not be modified.

## Reuse map

| Area | Reuse now | PPTX-specific work |
| --- | --- | --- |
| OPC | Use `ooxml-opc` package limits, guarded entry reads, relationships, and save primitives. | Presentation part graph and content-type validation. |
| DrawingML | Move conservative, format-neutral color, fill, outline, transform, theme, geometry, chart, and picture primitives into `ooxml-drawingml`; retarget DOCX parse/layout without changing wire shapes. | PresentationML shape-tree traversal, placeholder semantics, and master/layout inheritance. |
| Parsing/saving | Follow `docx-parse` bounded quick-xml structure and S13 untouched-part preservation. | Presentation, master, layout, slide, notes-independent text-body, media, and relationship models. |
| Editing | Copy the `docx-edit` pilcrow-as-character story invariant, `EditCtx`, typed receipts, local-origin undo policy, and raw update/state-vector surface into `pptx-edit`. | Movable slide sequence, per-shape stories, shape geometry operations, and deck receipts. Deduplicate with DOCX after the parallel mission. |
| Text/layout | Use `ooxml-text` shaping and the DOCX display-list shape/glyph conventions. | Text-box wrapping/alignment/basic autofit, slide z-order, placeholder composition, and slide hit testing. |
| Wasm/package | Mirror `build-docx-wasm.ts`, external wasm loading, generated glue, and canvas replay boundaries; prefer one combined PPTX module. | Deck-oriented public handle and PPTX display-list decoder. |
| Collaboration | Port the XLSX protocol/provider verbatim and adapt only the replica binding. | PPTX session convergence and two-browser proof. |
| React/demo | Match the existing DOCX/XLSX package metadata, README style, and demo shell. | Slide strip, slide canvas, caret/selection, formatting toolbar, fixture deck, and `/pptx`. |

## Phases

1. Extract only proven shared DrawingML units. Keep DOCX JSON and rendering output unchanged; shrink the extraction if parity is uncertain.
2. Add `pptx-parse` with bounded quick-xml 0.41 parsing, no external fetches, guarded OPC reads, and raw untouched-part preservation. Generate and test a branded fixture with `scripts/create-demo-deck.ts`.
3. Add `pptx-edit` with a yrs deck root, slide ordering, one pilcrow story per text body, shape operations, typed receipts, local-only undo, v1 update exchange, convergence tests, and a wasm feature.
4. Make `pptx-render` consume the parsed/editable deck, shape text with `ooxml-text`, resolve master/layout placeholders, emit the shared display-list family, and expose shape/text hit tests.
5. Add ESM-only `@betteroffice/pptx` and `@betteroffice/pptx-react` 0.0.1 packages, generated wasm glue with ignored binaries, canvas replay, working text-box editing, slide controls, and fixed changeset grouping.
6. Add the ported collaboration module, protocol/provider/wasm convergence tests, and a two-context browser proof with screenshots.
7. Replace the PPTX placeholder demo, mark the format live, extend root builds, run production/browser checks, and prepare the scoped conventional PR.

## Phase gate

Each green checkpoint runs workspace tests, lint and formatting, wasm builds, affected TypeScript checks, the DOCX build fence, package tests, demo production build, committed-wasm checks, and dependency-policy checks. Commits use explicit paths, bypass repository hooks as requested, carry the Codex co-author trailer, and are pushed before the next phase begins.
