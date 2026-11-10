---
'@betteroffice/docx-native': patch
'@betteroffice/pptx-native': patch
'@betteroffice/pptx': patch
'@betteroffice/rust-crates': patch
---

EMF and WMF pictures that draw the same bitmap several times, or whose EMF+ drawing falls back to the GDI one, no longer exceed the bitmap budget: equal bitmaps are decoded and counted once, and a set-aside pass's bitmaps are released first.
