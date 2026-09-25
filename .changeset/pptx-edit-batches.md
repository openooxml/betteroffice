---
"@betteroffice/pptx": minor
"@betteroffice/pptx-react": minor
"@betteroffice/rust-crates": minor
"@betteroffice/python-pptx": minor
---

Add version-checked, all-or-nothing PPTX edit batches. `PresentationHandle.version()`, `readContent()` and `findText()` return slides and story text with the session version they were read at; `validateEdits()` and `applyEdits()` resolve every step against that version and either commit text insertion, replacement, deletion and formatting, paragraph alignment, speaker notes, and shape rectangle, fill and outline changes as one transaction and one undo step, or return a typed refusal (`stale-version`, `missing-target`, `ambiguous-target`, `content-mismatch`, `overlapping-steps`, `unsupported`, `invalid-step`, `limit-exceeded`) with the deck untouched. `history: "none"` keeps a batch out of undo history, and `source` records provenance only. `PptxEditorApi` gains the same operations, flushing pending input first, refreshing the editor once after an applied batch, and refusing with `read-only` while the editor is read-only. The Rust `DeckSession` and `Presentation` and the Python `Presentation` expose the same API, and proposal previews and acceptance now share the batch staging. Text steps stay within one paragraph and leave fields and soft line breaks whole; ids and versions are session-scoped.
