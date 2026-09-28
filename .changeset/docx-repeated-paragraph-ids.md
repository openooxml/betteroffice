---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-docx': patch
---

Paragraphs whose `w14:paraId` repeats an earlier paragraph's are addressable again through the native `Document` and the Python binding: each reads with a fresh ID that no part of the package uses, which `replace_text` and `replace_paragraph_text` edit and a save writes once that paragraph is edited or the model is changed through `model_mut`, while unedited paragraphs save with their authored IDs. A paragraph ID that still matches several paragraphs is refused with `Error::AmbiguousParagraph` (`KeyError` in Python) instead of editing the first. Paragraph IDs and comment companion references written with character references (`&#x44;`) are now read as the parser reads them, so IDs allocated by `persistParagraphIds()` and the native facade never collide with them and companion references follow a renamed comment paragraph.
