---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Committing a shown formula unchanged keeps it a formula when it references another workbook (`[1]Sheet1!A1`) or was stored with a leading `=`, which now shows one `=` and calculates. Typing into a Text-formatted cell still enters text, as in Excel.
