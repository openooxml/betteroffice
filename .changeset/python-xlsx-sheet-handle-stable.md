---
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': minor
---

A held `Sheet` keeps addressing the same sheet as sheets are added, moved, or removed. A handle to a removed sheet raises `KeyError`. The Rust `Workbook::sheet_ids` returns the sheet ids alone.
