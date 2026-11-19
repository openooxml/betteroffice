---
"@betteroffice/xlsx": minor
"@betteroffice/rust-crates": minor
"@betteroffice/xlsx-native": patch
"@betteroffice/python-xlsx": patch
"@betteroffice/agents": patch
"@betteroffice/xlsx-react": patch
---

XLSX rich text cells paint each run in its own size, style, colour and, where available, font on screen, in PNG exports and in print; formula, number-formatted, rotated, stacked and right-to-left cells stay plain. Rust code building or fully matching `DrawCmd::Text` must add `runs`.
