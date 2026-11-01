---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/agents': patch
'@betteroffice/rust-crates': patch
---

Formulas written in BetterOffice that use newer Excel functions (XLOOKUP, IFS, LET, LAMBDA, FILTER and others) are saved with the `_xlfn.` prefixes Excel expects, so Excel no longer opens them as #NAME?. Unedited formulas keep their exact text.
