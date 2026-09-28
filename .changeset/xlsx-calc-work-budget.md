---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: recalculation keeps a hostile workbook inside its work budget, so opening one returns in milliseconds instead of running for minutes: `MAKEARRAY`, `SEQUENCE` and the other array builtins charge the budget before they build a block, callbacks stop at the first refusal, and long callback bodies, repeated arguments, `SUMPRODUCT` operands, matrix products, `TEXTSPLIT` separators and nested array fallbacks pay for the work they do. A block of lookup keys pays for the values its searches visit, and `MAKEARRAY` charges its cells once, so a 600,000-cell block that fits the budget now evaluates instead of returning `#NUM!`.
