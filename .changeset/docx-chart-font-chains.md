---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/python-docx': patch
'@betteroffice/rust-crates': minor
---

Native DOCX rendering no longer fails on pages with charts or comma-separated font names. Chart text naming a theme font now uses the document theme's font in every engine.
