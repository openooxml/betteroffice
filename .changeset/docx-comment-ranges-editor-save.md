---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

`saveYrsDocx` and the editor's Save write comment ranges that start or end inside a hyperlink, including one that holds a field or an equation, or inside a tracked change, or that follow a field, content control, equation, table or page break in their paragraph, as a paired start and end with a reference over the same text, and reopen them over that text. The editor's Save now also writes the ranges of comments added in the editor, in table cells and content controls too.
