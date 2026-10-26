---
"@betteroffice/agents": minor
"@betteroffice/xlsx": patch
---

Adds opt-in `protectedSheets: "allow"` to agent cell writes, formatting, sorting and find/replace. Protection is kept, and `protectedCells` reports edits through protection. The default remains `"refuse"`.
