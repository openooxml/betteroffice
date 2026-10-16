---
"@betteroffice/docx": minor
---

`CollaborationProvider` streams large initial syncs as bounded Update frames (`syncChunkBytes`, default 1 MiB) and applies them as one change (`maxSyncBytes`, default 64 MiB). Replicas opt in with `encodeStateAsUpdateChunks` and `applyUpdateChunks`. Frame and buffer options now have upper bounds.
