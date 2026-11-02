---
"@betteroffice/xlsx": patch
"@betteroffice/pptx": patch
"@betteroffice/docx": patch
"@betteroffice/xlsx-native": patch
"@betteroffice/pptx-native": patch
"@betteroffice/docx-native": patch
---

Worker sessions now run on Node.js >= 20.16 / 22.3 through `worker_threads`, keeping the process alive until disposed. DOCX worker editing needs `OffscreenCanvas` and fails with a clear error without it (Node.js, Bun, Deno); open, read and save still work.
