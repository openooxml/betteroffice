---
"@betteroffice/xlsx-react": patch
"@betteroffice/xlsx": patch
---

Preserves untouched cells and queued edits when saving or recovering with opt-in `asyncApi`. Honors `onSaveRequest`, loads the async editor separately, and reports missing canvas support or legacy API access with directed errors, including `XlsxAsyncApiSyncAccessError`.
