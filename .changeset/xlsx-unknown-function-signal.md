---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

Formulas calling a function the engine does not know keep the value Excel saved, even inside IFERROR, and are listed as unevaluated with the function name. GAMMA, GAMMALN and FACT are exact at integers.
