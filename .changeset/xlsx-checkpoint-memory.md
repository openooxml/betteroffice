---
"@betteroffice/xlsx": patch
---

Recovery checkpoints of large workbooks with long undo history no longer run out of wasm memory: each distinct sheet state is written once. A checkpoint over the capture budget is refused, and the session keeps working. Engine panics are now logged to the console.
