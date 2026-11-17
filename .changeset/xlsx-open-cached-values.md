---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/rust-crates': minor
---

Adds the `openWorkbook` option `recalculate: 'stale'`, which shows the formula results a file cached and recalculates only when a formula has none or `calcPr` requests it. Defaults to `'all'`.
