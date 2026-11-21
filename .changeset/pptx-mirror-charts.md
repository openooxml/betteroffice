---
"@betteroffice/pptx": patch
"@betteroffice/pptx-native": patch
"@betteroffice/pptx-react": patch
"@betteroffice/rust-crates": minor
---

Adds an optional `data` to chart display-list primitives with categories and each series' values as its data labels format them. Screen reader slide mirrors read a chart as a figure holding its values as a table; charts without data keep their summary. Supply `Primitive::Chart`'s `data`.
