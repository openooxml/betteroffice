---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

In DOCX, the paragraph mark shown after a tracked inserted or deleted paragraph is drawn from the font's outline when the document already loads the default text font, so it looks the same on every render; otherwise it is drawn as before.
