---
'@betteroffice/rust-crates': patch
'@betteroffice/xlsx': patch
---

After two row or column edits in one save, a shared formula with a gap is no longer written back as a different formula. A group whose members moved unevenly, or fall outside the written range, is now written with each cell's own formula.
