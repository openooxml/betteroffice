---
"@betteroffice/docx-react": minor
"@betteroffice/xlsx-react": minor
---

Plugins gain experimental `selection-settle` events, `snapshot.selectionSettled` and `{ kind: 'selection' }` anchor geometry. DOCX plugins receive header and footer caret moves. Remote edits removing the caret's story clear selection. Supply `selectionSettled` in DOCX and XLSX snapshot providers and implement XLSX's new `readAnchorGeometries` in geometry providers.
