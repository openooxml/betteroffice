---
'@betteroffice/docx-react': patch
'@betteroffice/docx': patch
---

In the worker editor, copy now completes with the selection it was pressed on, even when a click, caret move, paste or table cell selection changes the selection right after. A cut whose selection changes before its read finishes is still cancelled.
