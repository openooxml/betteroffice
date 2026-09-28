---
'@betteroffice/rust-crates': patch
'@betteroffice/python-docx': patch
---

Paragraphs whose `w14:paraId` repeats an earlier paragraph's are addressable again through the native `Document` and the Python binding: each reads with a fresh ID that `replace_text` and `replace_paragraph_text` edit and a save writes once that paragraph is edited, while unedited paragraphs save with their authored IDs. A paragraph ID that still matches several paragraphs is refused with `Error::AmbiguousParagraph` (`KeyError` in Python) instead of editing the first.
