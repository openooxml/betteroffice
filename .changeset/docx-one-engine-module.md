---
"@betteroffice/docx": patch
---

The DOCX editor loads one engine wasm module instead of also fetching the layout module on the main thread and in its worker. Standalone font measurement now works beside an open editing session instead of throwing.
