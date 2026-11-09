---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/docx-native': patch
---

Clearing or changing character spacing, scaling, kerning or raised text, and clearing a theme font, no longer rebuilds the paragraph on save, so its hyperlinks, smart tags and drawings are kept as they were.
