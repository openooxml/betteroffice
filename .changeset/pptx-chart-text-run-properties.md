---
"@betteroffice/pptx": patch
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Draw chart titles, axis labels, legends and data labels in the family, slant and character spacing their `c:txPr` declares, instead of the theme minor font upright and untracked. A chart title's own `c:rich` run properties now override the paragraph default they sit under, and a reattached source refreshes the text properties of a stored chart. A DOCX chart paints the tracking its text properties declare instead of only reserving room for it.
