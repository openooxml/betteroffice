---
"@betteroffice/docx-react": minor
---

**Breaking:** In read-only and viewing sessions, `getEditorRef`, `getSelectionInfo` and `getPositionAtPoint` return `null`, formatting and comment edits are refused, and `onContentChange` never fires. Use `readSelectionInfo`, `readPositionAtPoint` and `onDocumentChange` instead.
