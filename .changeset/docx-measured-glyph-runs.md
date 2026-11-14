---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/python-docx': patch
'@betteroffice/rust-crates': minor
---

DOCX lines measured with authoritative shaping now carry the glyphs measurement shaped them with, as `glyphRuns`, packed into one string per line. Rendering does not use them yet.
