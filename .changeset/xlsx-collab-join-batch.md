---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
---

Worker editors joining a collaboration room merge queued remote updates in one worker command instead of replaying them one at a time. Custom async replicas can opt in through the optional `applyUpdateChunks`.
