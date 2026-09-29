---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Add `getMemoryStats()` and opt-in `onMemoryPressure` and `memoryBudget` for the editor's wasm memory. A worker that runs out of memory, or past the opt-in `memoryBudget.workerLimitBytes`, is replaced once, then reported as `ResidentWorkerOutOfMemoryError` without a main-thread fallback; other failures fall back as before.
