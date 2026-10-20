---
"@betteroffice/docx-react": minor
---

**Breaking:** `collaboration.onReplica` receives a replica whose synchronous `encodeStateVector`, `encodeStateAsUpdate` and `encodeStateAsUpdateChunks` throw; call them through `replica.async`. A pinned `clientId` whose clocks would be reused opens with a fresh id and a warning.
