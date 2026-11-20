---
"@betteroffice/docx": minor
"@betteroffice/docx-native": minor
"@betteroffice/rust-crates": minor
---

Unmodelled `w:` inline elements are kept on save. Parse warnings gain `DOCX_UNMODELED_INLINE` when such an element holds text, as tracked changes inside a hyperlink do. Rust: `RawOp` gains `RawOp::SetRawMark`, so code matching every variant must handle it.
