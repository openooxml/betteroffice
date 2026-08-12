---
"@betteroffice/rust-crates": patch
"@betteroffice/docx": patch
---

`nextColumn` section starts no longer stop a document from opening. `w:sectPr/w:type w:val="nextColumn"` (ST_SectionMark) parsed fine, but the layout engine's `SectionBreakType` had no matching variant, so the section start reached the layout request as a string serde could not deserialize and the whole `layoutDocumentWithRegionsJson` envelope failed — a document using it never rendered at all. Two quieter paths dropped the value before that: the layout bridge mapped it to `None`, and the Yrs seeding pass filtered `sectionStart` through an allowlist that omitted it, so the paragraph kept its `sectPr` but lost the `sectionBreakType` attribute the bridge prefers as a fallback.

`nextColumn` is now a first-class break type end to end. A section starting in the next column keeps the column band in force and moves to the next column, deferring its page geometry like a `continuous` break; out of the last column — so always in a single-column section — it falls through to a new page, matching Word and the existing column-break block. A break already sitting at the top of an empty page advances nowhere, so a leading `nextColumn` no longer emits a blank first page, the same idempotence a `nextPage` break has always had.

Word cannot change page size or column count part-way down a sheet, so a `nextColumn` section that differs in either is promoted to a page break; a section that only changes the column gap keeps the old band to the end of the sheet and its `w:cols` take effect on the next page, which the paginator now queues alongside the page size and margins it already deferred.
