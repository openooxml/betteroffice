---
"@betteroffice/docx-react": patch
---

`DocxEditorRef` adds async twins of its synchronous document members, such as `readDocument`, `readSelectionInfo` and `findParagraphs`, plus `onDocumentChange`; plugin geometry adds `readPositionAtPoint` and `readAnchorGeometry`. The synchronous members and the `onChange` prop are deprecated.
