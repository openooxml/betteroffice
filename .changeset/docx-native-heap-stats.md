---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/docx-react': patch
---

The native DOCX engine reports its editing heap per thread, listed by `wasmModuleMemories()`. Pressure events on native need a host memory budget; its heap limit is not enforced yet.
