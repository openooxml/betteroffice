---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
---

Hang each paragraph-anchored wrap band off its own anchor paragraph, and give a bare section-break paragraph mark no line. Bands whose `wp:positionV` is `relativeFrom="paragraph"` were grouped with nearby anchors and re-based onto the earliest one, so a band belonging to a later paragraph narrowed lines above it and forced wraps Word does not make; each band now starts at its own anchor's place in the flow and survives until the flow passes it. A `w:p` whose `w:pPr` carries a `w:sectPr` and holds no runs is the section break itself: Word prints no line for it, measured at 0.0pt against Word's own PDF export in three documents, where the same paragraph mark with run content still keeps its line.
