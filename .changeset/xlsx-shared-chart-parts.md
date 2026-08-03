---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

A chart or drawing part that two sheets both anchor is no longer written from whichever sheet came last. Saving such a workbook used to patch the shared part once, silently rewriting the chart the other sheet shows: inserting a row on one sheet moved its references and, on reopen, the other sheet's chart had moved with it. A part is now written only when every sheet anchoring it patches it into the same bytes, which covers the sheet a cache is rebuilt against, and the save is refused with an error naming the part and both sheets otherwise. Workbooks whose sheets want the same content out of a shared part — including every workbook that shares no part at all — save exactly as before, byte for byte.
