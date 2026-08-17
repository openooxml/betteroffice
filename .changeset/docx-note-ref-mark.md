---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Edited footnotes and endnotes keep their `w:footnoteRef` / `w:endnoteRef` mark on save, so Word still shows the note number in front of the note text. A footnote or endnote reference now counts as one position everywhere, so bookmarks and comments around a reference whose id has more than one digit no longer shift or disappear.
