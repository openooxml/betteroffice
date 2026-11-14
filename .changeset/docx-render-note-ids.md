---
'@betteroffice/rust-crates': patch
'@betteroffice/python-docx': patch
---

DOCX `render_png` renders documents with footnotes or endnotes again. The layout request now writes integral numbers as integers, so note ids, the note numbering start and footnote columns reach layout in the form it reads.
