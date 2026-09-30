---
"@betteroffice/docx": minor
---

**BREAKING (direct wasm consumers only):** `apply_delete`/`apply_delete_profiled` take a third `count` argument; pass `1` for the previous behaviour. `YrsSession.applyDelete` is unchanged.
