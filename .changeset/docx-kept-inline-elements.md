---
"@betteroffice/docx": minor
"@betteroffice/docx-native": minor
"@betteroffice/rust-crates": minor
---

Unmodelled `w:` inline elements are kept on save. Parse warnings gain `DOCX_UNMODELED_INLINE` for kept ones holding text, such as tracked changes inside a hyperlink, which an edit of their paragraph still drops. Rust: `RawOp` gains `RawOp::SetRawMark`, so exhaustive matches must handle it.
