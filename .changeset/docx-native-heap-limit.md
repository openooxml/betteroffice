---
'@betteroffice/docx-native': patch
'@betteroffice/docx': patch
'@betteroffice/docx-react': patch
---

The native DOCX engine now enforces the worker heap limit: a call past it fails with the same out-of-memory error as wasm, and its object refuses later calls. A limit turns heap counting on; without a limit or `startHeapCounting()`, counting costs nothing.
