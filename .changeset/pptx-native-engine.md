---
'@betteroffice/pptx': patch
---

By default, `@betteroffice/pptx` uses the native engine in Node, Bun and Deno when a matching `@betteroffice/pptx-native` is installed, falling back to wasm otherwise. `engineBackend()` reports the engine; `BETTEROFFICE_ENGINE=wasm|native` overrides selection.
