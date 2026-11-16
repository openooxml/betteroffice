---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': minor
---

Excel tables now render their built-in or workbook-defined table style by default: header, total, banded and first or last column fills, font colours and bold. Breaking: Rust `Table` gains a public `style` field, so struct literals must set it.
