---
"@betteroffice/pptx": patch
"@betteroffice/xlsx": patch
---

`initWasm()` without an input now works on Node.js: the packaged WASM asset is read from disk instead of fetched as a `file:` URL. When the asset cannot load, the error names it and says to pass bytes to `initWasm`.
