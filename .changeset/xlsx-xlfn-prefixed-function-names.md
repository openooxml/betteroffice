---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
---

Resolve the `_xlfn.` prefix Excel stores on post-2007 functions. Excel writes every function added after the 2007 grammar under `_xlfn.`, and the worksheet-only subset under `_xlfn._xlws.`, so a stored `_xlfn.TEXTJOIN` never matched the builtin table and recalculated to `#NAME?` even though TEXTJOIN is implemented. The prefix is now stripped before the lookup, so `TEXTJOIN`, `CONCAT`, `IFNA`, `IFS`, `SWITCH`, `XLOOKUP`, `RANK.EQ`, `STDEV.P` and `VAR.S` compute under their stored names; a prefixed function that is genuinely not implemented, such as `_xlfn._xlws.FILTER`, still reports `#NAME?`. 25 of the 208 xlsx fidelity samples store such a name. Most already displayed correctly because recalculation keeps Excel's cached value when it meets a function it does not know, so the change is visible only where that cached value was absent: sample sheetpedia-517dc2896b7e 0.6141 → 0.6147, sheetpedia-f155b9332371 0.9177 → 0.9178, sheetpedia-9eaddd49cfb3 0.9162 → 0.9162, sheetpedia-6f3940c8975e 0.7614 → 0.7615, sheetpedia-06dfb94ff719 0.5785 → 0.5785.
