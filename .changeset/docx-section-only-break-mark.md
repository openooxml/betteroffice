---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
---

Give a section that holds nothing but its own break mark the line Word gives it. A `w:p` whose `w:pPr` carries a `w:sectPr` and holds no runs is the section break itself and still prints no line, but when that mark is all its section contains — the whole section is the break — Word lays the section out one line tall and every later line sits below it. `oxi-en-correspondence-03` opens with such a section; its body now starts where Word's PDF export starts it, 25px lower at 150dpi, and its penalized SSIM rises from 0.6464 to 0.7977.
