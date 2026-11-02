---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

Calculation matches Excel at numeric limits: numbers become text with 15 significant digits (`""&(0.1+0.2)` is "0.3"), formula literals keep 15 digits within Excel's range, subnormal results are 0, and MOD, MULTINOMIAL, BESSELK, the TBILL, T and F functions follow Excel's limits.
