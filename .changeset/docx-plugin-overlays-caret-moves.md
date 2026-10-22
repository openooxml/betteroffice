---
"@betteroffice/docx-react": patch
---

Plugin overlays in the editable editor no longer unmount and remount on caret moves. Plugin geometry keeps answering from the last built frame while the document version and layout are unchanged, and still answers `layout-unavailable` after an edit until it is rebuilt.
