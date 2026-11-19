---
"@betteroffice/xlsx": minor
"@betteroffice/rust-crates": minor
"@betteroffice/xlsx-native": patch
"@betteroffice/python-xlsx": patch
"@betteroffice/agents": patch
"@betteroffice/xlsx-react": patch
---

XLSX rich text cells paint each run in its own size, style, colour and available font on screen, in PNG and print. Formula, number-formatted, rotated, stacked, right-to-left cells and cells without room stay plain. Rust code building or fully matching `DrawCmd::Text` must add `runs`.
