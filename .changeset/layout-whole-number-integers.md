---
'@betteroffice/rust-crates': patch
'@betteroffice/python-docx': patch
---

Layout requests accept a whole-number float such as `1.0` where an integer is expected, as a parsed document forwarded from Python or Rust writes it, so documents with footnotes or endnotes lay out. Fractional values are rejected, naming the field.
