---
"@betteroffice/docx": minor
---

Collaboration adds experimental replicas whose state lives in a worker (`createAsyncCollaborationReplica`, `AsyncReplicaAdapter`) and opt-in chunked initial sync (`syncChunkBytes`, `maxSyncBytes`). Offline edits converge after reconnecting, and updates over 64 MiB are refused.
