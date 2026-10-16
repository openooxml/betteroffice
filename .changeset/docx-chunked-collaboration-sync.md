---
"@betteroffice/docx": minor
---

`CollaborationProvider` can stream large initial syncs as bounded Update frames that the joiner applies as one change (`maxSyncBytes`, default 64 MiB). Chunked answers are opt-in for point-to-point transports (`syncChunkBytes`). Replicas opt in with `encodeStateAsUpdateChunks` and `applyUpdateChunks`. Frame and buffer options now have upper bounds.
