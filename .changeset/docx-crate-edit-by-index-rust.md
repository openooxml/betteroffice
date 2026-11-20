---
'@betteroffice/rust-crates': minor
---

Adds `Document::replace_paragraph_text_at(index, text)`, which edits a body paragraph by its position in `paragraphs()`, including paragraphs without a `w14:paraId`. Untouched content keeps its bytes, and the edit writes no paragraph ID.
