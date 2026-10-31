---
"@betteroffice/fonts": patch
---

Packed faces fall back to the decoder shipped with the package when `<baseUrl>/decoder/fontpack.wasm` is missing, so a self-hosted `baseUrl` needs only the font assets. When no decoder loads, the error names every location tried.
