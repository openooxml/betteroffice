---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: whole-column and whole-row references count the blanks past the used range again, so `COUNTBLANK(A:A)`, `COUNTIF(A:A,"")`, `COUNTIF(A:A,"<>x")` and `TEXTJOIN(",",FALSE,A:A)` match Excel and `VLOOKUP`/`HLOOKUP` can return a column or row past the data. Blocks computed from them keep those blanks too, in array formulas and `LET` as well: `SUMPRODUCT(--(A:A=""))`, `COUNT(1/(A:A=""))`, `MATCH(TRUE,A:A="",0)` and `UNIQUE(A:A)` answer as Excel does, and whole references on different sheets or behind names pair row for row in `SUMIF`, `SUMIFS`, `AVERAGEIF`, `CORREL`, `SUMPRODUCT`, `XLOOKUP`, `FILTER` and array arithmetic. `SUMPRODUCT` requires operands of one shape and reads logicals as zero in array formulas as it already did elsewhere, and `SUMPRODUCT` and `MMULT` over whole references return a value or `#VALUE!` instead of panicking.
