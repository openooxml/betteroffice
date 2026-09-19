---
'@betteroffice/docx': patch
---

Keep a text-wrapping anchored image on its page. An anchor offset that pushed the image past the bottom edge used to place it off the sheet, where it rendered nowhere and the picture was simply lost; Word clamps such a float so its bottom sits on the page edge.
