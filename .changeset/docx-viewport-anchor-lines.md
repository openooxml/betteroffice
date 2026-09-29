---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Capture the viewport's scroll anchor from the lines of the pages on screen instead of every line in the document. After each keystroke the editor anchors the viewport to its topmost visible line, and finding it computed the visual lines of the whole document first, 40–95 ms per keystroke on a 267-page document. Display-list queries gain `visualLinesOnPage`, the anchor reads only the pages around the viewport, and grouping a page's lines no longer searches every line already found on it.
