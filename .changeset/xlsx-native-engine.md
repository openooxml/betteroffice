---
'@betteroffice/xlsx': patch
---

Node, Bun and Deno use the native engine when @betteroffice/xlsx-native is installed (same API; browsers keep WebAssembly); engineBackend() reports which engine runs; BETTEROFFICE_ENGINE=wasm|native overrides.
