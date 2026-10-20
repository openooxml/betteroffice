---
"@betteroffice/xlsx-react": minor
---

**Breaking:** Synchronous API members such as `save`, `cell` and `editCell` throw `XlsxAsyncApiSyncAccessError`, which names the replacement: saving, undo and redo use `commands.execute`; cell reads, edits, selection and hit testing use the plugin `read`, `edits`, `navigation` and `geometry` APIs.
