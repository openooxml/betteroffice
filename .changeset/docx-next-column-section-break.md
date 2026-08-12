---
"@betteroffice/rust-crates": patch
"@betteroffice/docx": patch
---

`nextColumn` section starts no longer stop a document from opening. `w:sectPr/w:type w:val="nextColumn"` (ST_SectionMark, ECMA-376 §17.18.77) parsed fine, but the layout engine's `SectionBreakType` had no matching variant, so the section start reached the layout request as a string serde could not deserialize and the whole `layoutDocumentWithRegionsJson` envelope failed — a document using it never rendered at all. Two quieter paths dropped the value before that: the Yrs seeding pass filtered `sectionStart` through an allowlist that omitted it, and the layout bridge mapped it to `None`.

`nextColumn` is now a first-class break type end to end. A section starting in the next column keeps the column band in force, defers its page geometry like a `continuous` break, and moves to the next column; out of the last column — so always in a single-column section — it falls through to a new page, matching Word and the existing column-break block. The new section's `w:cols` apply where a fresh column region actually begins, and terminal column balancing no longer runs for a break that stayed inside its band.
