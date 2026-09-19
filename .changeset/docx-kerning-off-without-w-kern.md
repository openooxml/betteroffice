---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Stop applying OpenType pair kerning to runs that carry no `w:kern`. Word kerns only above a nonzero half-point threshold, but an absent threshold reached rustybuzz as "no feature overrides", which enables GPOS kerning by default and measured every line slightly narrow — enough to pull an extra word onto a line and shift the rest of a page. The measurement gate and the painted glyph pen now read the same rule, so both projections agree, and a run that sets `w:kern` at or below its font size still kerns. Shaping also restores visual glyph order for kerning-off RTL runs, which rustybuzz leaves in logical order when it routes them through the legacy `kern` table.
