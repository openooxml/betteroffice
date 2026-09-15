---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Refuse shape adjustment edits on custom geometry instead of silently dropping them on save, and report an exhausted shape id space instead of overflowing it.
