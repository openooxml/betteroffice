---
"@betteroffice/docx-native": patch
"@betteroffice/xlsx-native": patch
"@betteroffice/pptx-native": patch
---

Native engine classes now have `[Symbol.dispose]`, the same function as `free()`, as on the wasm packages, so `using` declarations release them at scope exit.
