---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Edited saves keep each cell's original style entry when its format is unchanged, so cell protection and rotation survive and unrelated sheets are written unchanged.
