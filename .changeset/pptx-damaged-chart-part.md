---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

A deck whose chart part cannot be read now opens with that chart missing, instead of the whole file being refused; slides, masters and text come through intact. Chart parts also get their own parse budget, so a valid deck whose charts carry large cached series opens with every chart, and the part the parser declined is still written back untouched on save.
