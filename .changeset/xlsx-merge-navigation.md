---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
---

Navigation and typing treat a merged cell as one cell, as Excel does: Enter, Tab, arrows, paging and jumps land on a merge's top-left cell and step past it, and typing into a merge edits that cell. Adds optional `SelectionData.merge` and a `'merge'` lookup direction.
