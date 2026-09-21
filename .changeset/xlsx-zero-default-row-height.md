---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
---

Print sheets that declare a zero default row height. A sheet that hides every unsized row writes `zeroHeight="1"` next to `defaultRowHeight="0"`, and the print-metrics validator required a default row height of at least one point, so such a workbook reported `viewport must have finite positive dimensions` and never rendered at all. Zero is now accepted, exactly as a zero default column width already was; rows carrying an explicit `ht` keep their height and unsized rows collapse. Sample sheetpedia-738df30d7382 went from failing to capture to 0.7516; no other sample in the 208-sample xlsx fidelity corpus changed by a single bit.
