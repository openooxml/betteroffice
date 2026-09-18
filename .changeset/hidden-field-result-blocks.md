---
'@betteroffice/docx': patch
---

Hide the cached result of a block-spanning suppressed field in documents that carry no `w14:paraId`. The result blocks were matched by paragraph id, which is optional in OOXML, so a field whose result spanned blocks kept rendering them while its inline display text was already blanked. Seeding now binds the field to the story blocks its cached result duplicates — tables and block content controls as well as paragraphs — so editing the story cannot shift the suppressed range.
