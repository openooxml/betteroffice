---
'@betteroffice/docx': patch
'@betteroffice/docx-react': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

In DOCX, the mark after a tracked inserted or deleted paragraph is drawn from the outline of the default text font when the document uses that font, so it looks the same on every render. Otherwise it is drawn as before.
