---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-react': minor
'@betteroffice/python-xlsx': minor
'@betteroffice/rust-crates': minor
---

**Breaking:** collaboration updates now carry workbook schema 7, which stores each array anchor's definition in one payload with its formula and keeps array results out of the shared document, and a schema 3–6 update is migrated to 7 when it is opened. 0.1.x and 0.2.x clients cannot open a collaboration document this version seeded, saved or migrated, including rooms whose workbook has no array formulas, so mixed-version collaboration across this upgrade is not supported: upgrade every client and service that reads or writes shared collaboration state together. Ordinary `.xlsx` files are not affected. A stored update nobody edited drops its arrays' cached results on migration, apart from arrays the engine cannot evaluate, which keep them as stored cells; one that was edited keeps them as the constants it stored them as, so those arrays show `#SPILL!` until the constants are cleared. In Rust, `CellState` gains an `array` field carrying an anchor's `ArrayDefinition`, and `Sheet::set_cell` now makes a cell that held an array's result the author's.
