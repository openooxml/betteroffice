---
"@betteroffice/docx": patch
"@betteroffice/docx-native": patch
---

Adds opt-in `fontPreflight`, returning DOCX font faces, script keys, and scan completeness without the editing engine. Its standalone module loads only when called or explicitly loaded with `preloadPreflightWasm`.
