---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
---

Documents whose XML parts start with a byte order mark and have bookmarks between paragraphs now save. Before, saving them failed with a malformed XML error.
