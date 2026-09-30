---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

With `experimentalWorkerOpen`, a read-only editor now serves host proposals, plugin overlays and navigation from the worker without loading its main-thread copy of the document. New `getParagraphIdentities` and `resolveParagraphAnchors` on the editor ref.
