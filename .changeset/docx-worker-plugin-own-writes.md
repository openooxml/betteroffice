---
"@betteroffice/docx-react": patch
---

In the worker-only editor, a plugin's `applyEdits` makes its batch current before resolving, so `setState(..., applied.version)` succeeds and the write no longer aborts the plugin's own load or change hook. Reads and navigation raced by an edit or a closed document refuse instead of rejecting.
