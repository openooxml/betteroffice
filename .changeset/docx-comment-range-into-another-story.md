---
'@betteroffice/docx': patch
---

A comment that ends in a different table cell, content control or text box keeps its range on save when the saved file holds it whole: one start, one end, then its reference. Otherwise it is saved as before, cut at its first paragraph's end.
