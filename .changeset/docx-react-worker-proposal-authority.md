---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

With `experimentalWorkerOpen`, read-only editors serve proposals, overlays, navigation and `search()` from the worker. New ref methods `getParagraphIdentities` and `resolveParagraphAnchors`; synchronous members that need the document throw `DocxReplicaNotReadyError` until it loads.
