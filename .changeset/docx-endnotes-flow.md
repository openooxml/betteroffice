---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/python-docx': patch
'@betteroffice/rust-crates': minor
---

Endnotes now follow the last body text, as in Word, and continue onto new pages under a continuation separator. Before, long endnotes pushed body text onto extra pages and were painted over it. Set new fields in `NoteAreaContract`, `NoteRegion`, `ReservedBlock` and `NoteLayoutItemContract` literals.
