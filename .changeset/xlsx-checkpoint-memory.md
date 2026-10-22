---
"@betteroffice/xlsx": patch
---

Recovery checkpoints of large workbooks with a long undo history no longer exhaust wasm memory: each distinct preserved sheet state is written once instead of twice per undo step. A checkpoint that would still exceed the capture budget is refused like any other failed capture, so the session and its recovery log stay alive. Engine panics are logged to the console instead of surfacing only as `unreachable`.
