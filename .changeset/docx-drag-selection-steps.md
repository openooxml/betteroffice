---
'@betteroffice/docx': minor
'@betteroffice/docx-react': patch
---

Drag selection does less work per step: the worker keeps parsed pages between steps and answers a step with one message. `createDisplayListQueries` takes an optional `storePages` cap on parsed pages.
