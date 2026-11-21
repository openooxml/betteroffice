---
'@betteroffice/docx-native': minor
---

Replaces the native facade with the same entry points and API as `@betteroffice/docx`, backed by the native engine by default. Import `@betteroffice/docx-native` for Node.js, Bun and Deno without WebAssembly initialization or assets. Use `parseDocx` in place of `openDocument`, `Document` and `DocxDocument`, and drop imports of the removed facade types `DocumentStructure`, `EditReceipt`, `FontFace`, `ImageResource`, `LayoutResult`, `OpenDocumentOptions`, `RenderedPage` and `SaveDocumentOptions`.
