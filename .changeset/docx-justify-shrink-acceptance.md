---
'@betteroffice/rust-crates': patch
'@betteroffice/docx': patch
---

Justified paragraphs break more like Word. A word that fits only by compressing the spaces now stays on the line only when its overflow is at most half the stretch the line needs without it. Otherwise it moves to the next line.
