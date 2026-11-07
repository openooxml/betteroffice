---
'@betteroffice/pptx-native': patch
'@betteroffice/pptx': patch
---

Native pptx slide PNGs reject with `picture-unavailable` when a picture cannot be drawn, as native docx pages do, instead of leaving it out.
