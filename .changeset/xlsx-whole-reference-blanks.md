---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: whole-column and whole-row references count the blanks past the used range again, so `COUNTBLANK(A:A)`, `COUNTIF(A:A,"")`, `COUNTIF(A:A,"<>x")` and `TEXTJOIN(",",FALSE,A:A)` match Excel and `VLOOKUP`/`HLOOKUP` can return a column or row past the data. Whole references on two sheets stay aligned in `SUMIF`, `SUMIFS`, `AVERAGEIF`, `CORREL` and `SUMPRODUCT`, and `SUMPRODUCT` and `MMULT` over them return a value or `#VALUE!` instead of panicking.
