---
'@betteroffice/rust-crates': patch
'@betteroffice/xlsx': patch
---

Inserting or deleting XLSX rows or columns now moves pictures and shapes with their cells on save, as charts already move. An edit that would leave one where a save cannot place it, such as an object wholly inside deleted rows, is refused.
