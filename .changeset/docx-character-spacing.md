---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
---

DOCX character spacing now applies after every character, including the last character of a word or run, as Word does. Spaced text measures, wraps and paints with the same advances, and carets, selections and hit tests follow the painted glyphs.
