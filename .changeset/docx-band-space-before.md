---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Test a paragraph's first line against full-width wrap bands from the paragraph top rather than the line top, so `w:spacing w:before` is covered by the test and spent again below whatever the line clears. Word probed at twip resolution: a `topAndBottom` band reaching into `[0, before + lineHeight)` moves the line to the band's wrap bottom and re-applies space-before there, at every band height from 0.5pt to 200pt, while a band starting exactly at the line's bottom leaves it alone. `w:distT` participates in the trigger and `w:distB` extends the landing. Lines after the first are tested at their own box and land on the wrap bottom with no extra space. Table cells now use the same paragraph-top origin as the body flow instead of a re-measure loop.
