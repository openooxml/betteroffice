---
'@betteroffice/pptx': patch
'@betteroffice/pptx-native': patch
'@betteroffice/python-pptx': patch
'@betteroffice/rust-crates': patch
---

Slide text stored with CR LF or a lone CR reads as one line break, as in PowerPoint, so an edited paragraph no longer saves a stray carriage return and the text can be written back unchanged.
