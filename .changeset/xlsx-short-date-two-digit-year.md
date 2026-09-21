---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
---

Print built-in number formats 14 and 22 with the two-digit year Excel gives them. Both were mapped to a four-digit year (`m/d/yyyy` and `m/d/yyyy h:mm`); ECMA-376 §18.8.30 gives each a two-digit one, and Excel 16.112.3 prints `1/1/22` and `1/10/24 20:37`, matched cell by cell against its own PDF exports over 1201 dates in 33 of the 208 xlsx fidelity samples with no counter-example. Format 14 is the most common format code in that corpus, carried by 43 samples. 34 samples move, none down: sample sheetpedia-d72939b08fd0 0.6327 → 0.6390, sheetpedia-441e60078db9 0.6331 → 0.6395, sheetpedia-26bb0bfaf865 0.5664 → 0.5727, sheetpedia-1c26b6e9518d 0.6629 → 0.6678, sheetpedia-9eaec72339d9 0.6647 → 0.6690, betteroffice-workbook 0.9004 → 0.9030. Together with the `_xlfn.` prefix fix on the same branch the corpus mean goes 0.7371 → 0.7375 over the 205 scored samples. The `Date` number-format mutation still maps to built-in 14, so its proposal preview now reads `7/1/26`.
