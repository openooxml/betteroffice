---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Saving a workbook now preserves the parts the model does not represent — charts, drawings, pivot tables, comments, macros, custom XML and their relationships — instead of rebuilding the package and dropping them. Sheets you did not touch are copied through byte for byte, so an edit on one sheet no longer strips hidden rows, outline levels, rich inline strings or shared-formula attributes from the rest of the workbook. The stylesheet is left alone unless styles actually change, and adding a format now patches one pool entry instead of rewriting every pool. Chartsheets keep their type, cells keep the shared-string entry they were authored against, freeze-pane and hyperlink edits reach both the worksheet and its relationship part, and an edited save drops the stale calculation chain so Excel recalculates on open.

Known limits. A sheet you edit is still reserialized from the model, so unmodeled row, column and cell markup on that one sheet is lost. Editing an existing style pool entry regenerates that entry from the modeled subset. Sheet rename, sheet removal and row or column edits are refused while a chart or pivot table is preserved, because their references cannot be rewritten; formulas naming a removed sheet keep that name instead of collapsing to `#REF!`. Collaborative sessions compare only the modeled workbook, so two peers holding the same cells but different charts or macros still accept each other as the same base.
