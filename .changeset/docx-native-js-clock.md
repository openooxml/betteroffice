---
"@betteroffice/docx-native": patch
---

Undo grouping reads JavaScript `Date.now`, as the wasm engine does, so edits group the same way on both engines, including under a mocked clock.
