---
"@betteroffice/docx-react": patch
---

`insertBreak({ type: 'page' })` on a paragraph with text now splits it and places the page break between the two parts, like the editor's own page break command, so the document keeps laying out.
