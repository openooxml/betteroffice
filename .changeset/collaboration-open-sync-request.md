---
"@betteroffice/docx": patch
"@betteroffice/xlsx": patch
"@betteroffice/pptx": patch
---

Deletions made offline now reach every peer after a reconnect, even when newer edits reach those peers first. Pinned client ids must now stay below `Number.MAX_SAFE_INTEGER - 1`.
