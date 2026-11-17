---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/rust-crates': minor
---

Renaming a sheet also rewrites data-validation, extension conditional-format and sparkline formulas that reference it, as Excel does; a rename they cannot follow is refused. Adds save-warning code `unrewritable-formula`.
