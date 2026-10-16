---
"@betteroffice/xlsx": patch
"@betteroffice/agents": patch
"@betteroffice/rust-crates": patch
---

Edit batches now write content into unlocked cells of protected sheets, and formats when the protection allows formatting cells. Locked cells stay refused. Values and formulas written into empty cells keep their row or column style, as typed input does.
