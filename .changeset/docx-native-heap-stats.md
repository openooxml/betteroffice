---
'@betteroffice/docx': minor
'@betteroffice/docx-native': minor
'@betteroffice/docx-react': patch
---

The native DOCX engine reports its editing heap per thread, listed by `wasmModuleMemories()` with `native: true`. Pressure events on native need a host memory budget; its heap limit is not enforced yet.
