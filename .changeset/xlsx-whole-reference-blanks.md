---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: whole-column and whole-row references count the blanks past the used range again, so `COUNTBLANK(A:A)`, `COUNTIF(A:A,"")`, `COUNTIFS` and `TEXTJOIN(",",FALSE,A:A)` match Excel and `VLOOKUP`/`HLOOKUP` can return a column or row past the data. `SUMIF`, `SUMIFS`, `AVERAGEIF` and the other criteria functions, `SUMPRODUCT` and `CORREL` pair whole references on different sheets, or with an anchored value cell such as `SUMIF(A:A,"",Sheet2!B1)`, row for row, and `SUMPRODUCT` and `MMULT` over whole references return a value or `#VALUE!` instead of panicking.
