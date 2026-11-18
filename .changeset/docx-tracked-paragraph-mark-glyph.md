---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

In DOCX, the paragraph mark shown after a tracked inserted or deleted paragraph is drawn from the font's outline like the rest of the text, not as browser text, so it looks the same on every render.
