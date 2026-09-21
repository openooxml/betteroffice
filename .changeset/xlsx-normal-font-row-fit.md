---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-xlsx': patch
---

xlsx: a row with no explicit height now fits the workbook's normal font. A cell that names no style used to be measured at a fixed 11pt while the text in it was drawn at the normal size, so a workbook whose normal font is Arial 10 or Calibri 12 got a row grid that disagreed with its own text. When printing, the row fit and the text now read the same resolved face, so the two cannot diverge.
