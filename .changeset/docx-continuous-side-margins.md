---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': minor
---

A continuous section that changes its left or right margins now uses them from the page where it starts, as Word does, so its text, tables and floating objects no longer overflow or sit at the old position. Rust literals must set `ParagraphFragment.horizontal_frame`, `TableMeasurementFrame.horizontal_frame` and `TableMeasurementFrame.page_width`.
