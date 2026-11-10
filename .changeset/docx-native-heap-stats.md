---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
---

The native DOCX engine now reports its editing heap per thread: live, peak and failed allocation bytes, and `wasmModuleMemories()` lists its `docx-edit` entry. Its heap limit is not enforced yet.
