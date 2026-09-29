---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Relayout documents with footnotes or floating tables incrementally after an edit. The final pagination pass, which carries the footnote reservations, now resumes from the retained pass instead of starting over, so only the pages an edit changes are rebuilt and re-sent; the reservation-free pass no longer replaces the retained state, and the measured blocks are fingerprinted once instead of twice. In a document with paragraph-anchored floats only the float flow segments (runs of blocks between page, column and section breaks) that hold a changed block are measured again. Layout results are unchanged.
