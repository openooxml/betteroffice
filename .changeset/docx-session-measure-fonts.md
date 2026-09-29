---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Give each editing session its own measurement font store. Two editors mounted on one page no longer share font ids, so opening a second document with embedded fonts no longer repaints the first one with the second one's glyphs.
