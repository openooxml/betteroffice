---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': minor
---

`parse_docx_s4` no longer fails on an unreadable relationships part other than the document's own. It reports a warning and leaves that part's references unresolved.
