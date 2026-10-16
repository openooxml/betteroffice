---
"@betteroffice/xlsx": patch
"@betteroffice/pptx": patch
---

Reconnecting replicas mark their opening sync request, and every peer answers it with one request of its own, so offline deletions reach peers that already received later edits. Pinned client ids must now stay below `Number.MAX_SAFE_INTEGER - 1`.
