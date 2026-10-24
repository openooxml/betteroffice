---
"@betteroffice/xlsx-react": minor
---

**Breaking:** Reading synchronous API members such as `saveAsync`, `cell` and `editCell` throws `XlsxAsyncApiSyncAccessError`, which names the replacement: `save()`, `readCells` and `applyEdits` on the API, `commands.execute` for undo and redo, and the plugin `navigation` and `geometry` APIs for selection and hit testing.
