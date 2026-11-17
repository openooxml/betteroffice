---
'@betteroffice/docx-react': patch
'@betteroffice/docx': minor
---

The clipboardSlice document read gains an optional cells range; omitting it keeps the old behaviour. In the worker editor, copy now completes with the selection it was pressed on, even when the selection changes right after.
