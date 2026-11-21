---
'@betteroffice/docx-native': minor
---

`@betteroffice/docx-native` replaces its facade with `@betteroffice/docx`'s API and entry points, using native execution by default. Node.js, Bun and Deno imports need no WebAssembly initialization or assets. Replace `openDocument`, `Document` and `DocxDocument` with `parseDocx` and drop `DocumentStructure`, `EditReceipt`, `FontFace`, `ImageResource`, `LayoutResult`, `OpenDocumentOptions`, `RenderedPage` and `SaveDocumentOptions` imports.
