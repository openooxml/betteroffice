---
"@betteroffice/docx": minor
"@betteroffice/docx-react": patch
---

A superseded display-list query facade now answers from the live layout only within its own document line, never after the editor releases that document, and never across documents when no `line` is given. New `endDisplayListQueriesLine(line)` export from `@betteroffice/docx/layout/render`.
