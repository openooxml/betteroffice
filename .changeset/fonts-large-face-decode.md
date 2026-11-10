---
"@betteroffice/fonts": patch
"@betteroffice/docx": patch
"@betteroffice/pptx": patch
"@betteroffice/vsdx": patch
---

Large bundled font faces decode outside the PPTX worker's engine, and registering a font in WebAssembly copies its bytes once, which lowers worker memory for CJK presentations.
