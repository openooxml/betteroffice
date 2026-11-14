---
'@betteroffice/xlsx': patch
---

Cell text on the canvas paints at its face's own width. Linux Chromium rounded every glyph advance to a whole pixel, so digits ran up to 8% wider than the workbook's layout.
