---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Screen-reader links in DOCX only expose safe targets, whatever their source. Following one from the keyboard or a screen reader opens it in a new tab without an opener, and an internal link moves the caret to its bookmark.
