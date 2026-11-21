---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/docx-react': patch
'@betteroffice/rust-crates': minor
---

Unless hidden text is shown, DOCX rows marked `w:trPr/w:hidden` with entirely hidden content occupy no height and paint neither borders nor shading. Marked rows with visible content render. Fully hidden tables take no space. Cells retain document row numbers. Initialize new `TableRow` and `TableRowFormatting` fields.
