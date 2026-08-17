---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

A hard line break now starts a new line even when no measurement font resolves, so a title split by `<w:br/>` no longer collapses onto one baseline. Those fallback lines also carry the paragraph's `w:spacing` line rule instead of a fixed single-spaced box.
