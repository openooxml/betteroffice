---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Recover from resident worker crashes, WebAssembly traps, and unanswered requests so the editor can fall back to the main-thread engine. Reset retained worker frames and queries when switching engines so fresh main-thread frames render immediately.
