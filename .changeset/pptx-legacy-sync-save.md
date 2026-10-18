---
"@betteroffice/pptx-react": patch
---

`legacySyncApi` without `experimentalWorkerOpen` runs the in-thread engine again, so the sync `save()` returns the file's bytes. The `experimentalWorkerOpen={false}` warning now says to omit `legacySyncApi` as well to get the worker-owned editor.
