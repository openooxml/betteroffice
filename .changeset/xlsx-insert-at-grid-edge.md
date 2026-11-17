---
'@betteroffice/rust-crates': patch
'@betteroffice/xlsx': patch
---

An XLSX column insert no longer refuses when column formatting reaches the last column: formatting pushed past the edge is dropped, as Excel does; content there still refuses. A column inserted inside a uniform width or hidden run takes that run's width. Last-row heights likewise.
