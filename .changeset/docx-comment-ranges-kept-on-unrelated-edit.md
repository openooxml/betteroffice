---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

A comment whose range sits inside a tracked change, a hyperlink or a content control no longer loses that range when another paragraph is edited and the save keeps the source paragraph. Deleting such a comment removes its markers.
