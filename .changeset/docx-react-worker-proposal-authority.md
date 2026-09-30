---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

With `experimentalWorkerOpen`, a read-only editor serves host proposals, plugin overlays and navigation from the worker without loading the document on the main thread. New `getParagraphIdentities` and `resolveParagraphAnchors` on the editor ref. Meanwhile, synchronous ref members that need the document throw the new `DocxReplicaNotReadyError`.
