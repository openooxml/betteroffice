---
'@betteroffice/docx': patch
---

Measure a line carrying only an inline image at exactly the image's height. Word gives such a line no descent under the image, so the extra buffer pushed tall images onto a fresh page and left the page they came from nearly empty.
