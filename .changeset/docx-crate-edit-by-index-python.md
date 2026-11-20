---
'@betteroffice/python-docx': minor
---

Adds `Document.replace_text_at(index, text)` and its alias `replace_paragraph_text_at`, which edit a body paragraph by its position in `paragraphs()`, including paragraphs without a `w14:paraId`. Untouched content keeps its bytes, and the edit writes no paragraph ID. The `Edit` receipt gains `index`.
