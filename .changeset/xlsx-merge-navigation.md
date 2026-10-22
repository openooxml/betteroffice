---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
---

Keyboard navigation treats a merged cell as one cell, as Excel does: Enter, Tab, the arrow keys and paging step past a merge and land on its top-left cell, and typing into a merge edits that cell.
