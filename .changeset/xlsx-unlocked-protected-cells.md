---
"@betteroffice/xlsx": patch
"@betteroffice/agents": patch
"@betteroffice/rust-crates": patch
---

Edit batches now write content into unlocked cells of protected sheets, and formats when the sheet's protection allows formatting cells. Locked cells are still refused, naming the cell and sheet, and agent tools suggest unlocked cells instead.
