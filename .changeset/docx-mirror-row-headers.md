---
"@betteroffice/docx": patch
"@betteroffice/docx-native": patch
"@betteroffice/rust-crates": minor
---

Screen reader page mirrors read the first column of a table with "Repeat as header row" rows and first-column formatting as row headers, and data cells name their row header. Column headers are unchanged. Rust literals must set `TableCellRef.is_row_header`, `TableCell.merge_continuation` and `TableBlock.first_column`.
