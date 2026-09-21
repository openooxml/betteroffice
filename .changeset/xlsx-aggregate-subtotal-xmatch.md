---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: recalculation now supports `SUBTOTAL`, `AGGREGATE` and `XMATCH`. `SUBTOTAL` selects one of eleven aggregates by code, with the 100-series treated the same as the 1-series since manually hidden rows are not modelled. `AGGREGATE` adds `LARGE`, `SMALL`, `PERCENTILE` and `QUARTILE` on top, and its error-ignoring options recompute over the cells that are not errors. `XMATCH` supports exact, next-smaller and next-larger match modes and a reverse search.
