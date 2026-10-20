---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
"@betteroffice/rust-crates": patch
---

End (and Shift+End) in the xlsx editor goes to the last visible cell with data in the current row, falling back to the used range's last column on an empty row, instead of the sheet's last column.
