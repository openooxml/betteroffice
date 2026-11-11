---
'@betteroffice/rust-crates': minor
'@betteroffice/docx': minor
'@betteroffice/docx-native': minor
'@betteroffice/docx-react': patch
'@betteroffice/agents': patch
---

A table that crossing undos left unlayoutable is now skipped instead of breaking layout; table commands and saves refuse with a coded error until redo restores it or it is deleted. Both engines add `unavailable_tables_from`.
