---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

Restore persisted collaboration snapshots after the recalculation performed on open. Recalculation dirties saved formula caches without creating an authored edit; snapshot adoption now uses the authority's existing pristine-state check, while real local edits and incompatible workbook structures still prevent replacement.
