---
"@betteroffice/docx-react": minor
---

Worker-opened editors load their editing copy right after the first page appears and apply keys typed earlier in order. Until then, synchronous reads return `null` or `[]`, `scrollTo*Id` returns `false`, and synchronous edits throw `DocxReplicaNotReadyError`; await `flushPendingInput()` or use the async members.
