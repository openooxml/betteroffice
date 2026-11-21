---
'@betteroffice/pptx': patch
---

The PowerPoint editor's worker loads a smaller WebAssembly core (`pptx_wasm_lean_bg.wasm`, about 108 KB less gzip) without the one-shot `compileSlideJson` and `renderPptxMarkdownJson`. Self-hosted builds must serve both `.wasm` files. Public APIs keep the full core.
