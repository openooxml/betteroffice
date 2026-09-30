---
"@betteroffice/docx-react": patch
---

`DocxEditorRef` adds async twins of its synchronous document members, such as `readDocument`, `readSelectionInfo` and `findParagraphs`, and plugin geometry adds `readPositionAtPoint` and `readAnchorGeometry`. The synchronous members are deprecated.
