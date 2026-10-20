---
"@betteroffice/pptx-react": minor
---

**Breaking:** `collaboration.onReplica` receives an `AsyncCollaborationReplica` from `@betteroffice/pptx`. Migrate to its promise-returning `encodeStateVector`, `encodeStateAsUpdate` and `applyUpdate`.
