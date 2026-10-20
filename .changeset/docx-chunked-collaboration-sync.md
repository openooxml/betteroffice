---
"@betteroffice/docx": minor
---

`CollaborationProvider` can stream large initial syncs as bounded update frames that the joiner applies as one change (`maxSyncBytes`, default 64 MiB), opt-in for point-to-point transports with `syncChunkBytes` and for replicas through `encodeStateAsUpdateChunks` and `applyUpdateChunks`. Frame and buffer options now have upper bounds.
