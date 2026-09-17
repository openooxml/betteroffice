---
'@betteroffice/docx': patch
---

Hide the cached result of a block-spanning suppressed field in documents that carry no `w14:paraId`. The result paragraphs were matched by paragraph id, which is optional in OOXML, so a field whose result spanned paragraphs kept rendering them while its inline display text was already blanked; they are now recognised by their position after the field's own paragraph.
