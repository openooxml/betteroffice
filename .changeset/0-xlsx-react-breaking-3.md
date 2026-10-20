---
"@betteroffice/xlsx-react": minor
---

**Breaking:** `collaboration.onReplica` receives a replica whose `encodeStateVector` and `encodeStateAsUpdate` throw. Call `await api.collaboration.encodeStateVector()` and `await api.collaboration.encodeStateAsUpdate()` on the `XlsxAsyncEditorApi` from `onReady` instead.
