---
"@betteroffice/docx": minor
"@betteroffice/docx-native": minor
"@betteroffice/rust-crates": minor
---

Unmodelled `w:` inline elements are kept on save. Parse warnings gain `DOCX_UNMODELED_INLINE` for unmodelled `w:` inline elements holding text that the model keeps, for example tracked changes inside a hyperlink, which a save that does not rewrite the paragraph keeps and an edit of the paragraph still drops, as before. Rust: `RawOp` gains `RawOp::SetRawMark`, so code matching every variant must handle it.
