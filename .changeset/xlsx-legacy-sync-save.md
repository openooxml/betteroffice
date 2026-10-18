---
"@betteroffice/xlsx-react": patch
---

`legacySyncApi` without `experimentalWorkerOpen` runs the in-thread engine again, so the sync `save()` returns the file's bytes.
