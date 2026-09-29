---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Add `DocxEditorRef.getMemoryStats()`, `wasmModuleMemories()` and an opt-in `onMemoryPressure` callback with `memoryBudget` levels (75% and 90% of the 4 GiB wasm limit by default). A resident worker that runs out of memory, or past the opt-in `memoryBudget.workerLimitBytes`, is replaced once and then reported through `onError` as `ResidentWorkerOutOfMemoryError` instead of moving its work to the main thread; every other worker failure still falls back to the main thread as before.
