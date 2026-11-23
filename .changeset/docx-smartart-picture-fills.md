---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': minor
---

Picture fills in SmartArt now resolve against the diagram drawing's own relationships, so they draw instead of failing the page natively and falling back to a solid fill on canvas. Pass a media map or `None` before the context in `parse_smart_art_from_drawing`.
