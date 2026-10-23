---
"@betteroffice/xlsx-react": minor
---

**Breaking:** Reading synchronous API members such as `save`, `cell` and `editCell` throws `XlsxAsyncApiSyncAccessError`, which names the replacement: saving, undo and redo use `commands.execute`; cell reads, edits, selection and hit testing use the plugin `read`, `edits`, `navigation` and `geometry` APIs.
