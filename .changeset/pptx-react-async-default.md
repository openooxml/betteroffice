---
"@betteroffice/pptx-react": minor
---

`PptxEditor` now defaults to the worker-only runtime with an async `onReady` API. `legacySyncApi` retains the hybrid sync API as a deprecated opt-in removed in 0.6.0; `experimentalWorkerOpen={false}` remains a deprecated in-thread opt-in.
