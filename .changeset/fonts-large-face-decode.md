---
"@betteroffice/fonts": patch
"@betteroffice/docx": patch
"@betteroffice/pptx": patch
"@betteroffice/xlsx": patch
"@betteroffice/vsdx": patch
---

Large bundled font faces decode outside the engine, and registering a font in WebAssembly copies its bytes once, which lowers worker memory for CJK documents.
