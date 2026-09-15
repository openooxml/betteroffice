---
"@betteroffice/xlsx": patch
"@betteroffice/python-xlsx": patch
"@betteroffice/rust-crates": patch
---

Load Excel shared formulas by expanding followers into plain cells with correct absolute and relative references, so they evaluate and save with correct values. Shared-formula markup is not written back: an edited sheet writes each follower as its own formula.
