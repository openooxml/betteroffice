---
"@betteroffice/docx-react": minor
"@betteroffice/xlsx-react": minor
---

Adds experimental `selection-settle` events, `snapshot.selectionSettled` and `{ kind: 'selection' }` anchor geometry for plugins; XLSX adds `readAnchorGeometries`. DOCX plugins now receive header and footer caret moves, and a remote edit removing the story that holds the caret clears the selection.
