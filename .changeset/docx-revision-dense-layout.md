---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

DOCX layouts dense with tracked changes no longer collapse when measurement fonts are unavailable; overflowing fallback text is truncated instead of overprinting following runs.
