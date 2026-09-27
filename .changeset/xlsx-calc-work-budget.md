---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: recalculation keeps a hostile workbook inside its work budget, so opening one returns in milliseconds instead of running for minutes: `MAKEARRAY`, `SEQUENCE` and the other array builtins charge the budget before they build a block, callbacks stop at the first refusal, and long callback bodies, matrix products, broadcasts and nested array fallbacks pay for the work they do. `MAKEARRAY` charges its cells once, so a 600,000-cell block that fits the budget now evaluates instead of returning `#NUM!`.
