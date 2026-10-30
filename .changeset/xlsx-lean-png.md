---
"@betteroffice/xlsx": patch
---

Adds opt-in `initWasm(input, { png: false })` and worker session `png: false` to load a smaller WASM binary without PNG export. PNG export remains enabled by default and on the native backend.
