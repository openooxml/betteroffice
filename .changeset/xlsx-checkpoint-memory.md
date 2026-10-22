---
"@betteroffice/xlsx": patch
---

Recovery checkpoints of large workbooks with long undo history no longer run out of wasm memory: each sheet state is written once. Checkpoints over the capture budget are refused while the session keeps working. Engine panics are logged.
