---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-react': minor
'@betteroffice/python-xlsx': minor
'@betteroffice/rust-crates': minor
---

**Breaking:** collaboration updates now carry workbook schema 7, which stores each array anchor's definition in one payload with its formula and keeps array results out of the shared document, and a schema 3–6 update is migrated to 7 when it is opened. 0.1.x and 0.2.x clients cannot open a document this version saved or migrated, so mixed-version collaboration across this upgrade is not supported: upgrade every client that shares a stored update. A stored update nobody edited drops its arrays' cached results on migration; one that was edited keeps them as the constants it stored them as, so those arrays show `#SPILL!` until the constants are cleared. In Rust, `CellState` gains an `array` field carrying an anchor's `ArrayDefinition`, and `Sheet::set_cell` now makes a cell that held an array's result the author's.
