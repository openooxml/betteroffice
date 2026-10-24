---
"@betteroffice/docx": patch
---

The document worker file is now `docxWorker.mjs` (was `residentEngineWorker.mjs`) and its thread is named `betteroffice-docx-worker`. Lazy chunks use hash-only names. Bundlers that copy or allowlist the worker file by name need the new name.
