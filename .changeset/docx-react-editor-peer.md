---
"@betteroffice/docx-react": patch
---

Editors opened with the document worker load their editing copy right after the first page appears, and keys typed before it is ready are applied in order with normal undo grouping. While it loads, synchronous reads return `null` or `[]`, `scrollTo*Id` returns `false`, and synchronous edits such as `proposeChange` throw `DocxReplicaNotReadyError`; await `flushPendingInput()` or use the async members.
