---
'@betteroffice/docx': patch
---

Generated keys and nonces on the wasm build use the platform's random generator when the host provides one, and the full 64-bit range either way, as the native build does.
