---
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-native': minor
'@betteroffice/python-xlsx': minor
'@betteroffice/rust-crates': patch
---

Adds `saveWithWarnings()` (`save_with_warnings()` in Python): the saved bytes plus `{ code, part, element?, sheet?, message }` for each part a save kept as the file wrote it. Plain `save()` reports them as Node or Python warnings; ordinary saves carry none.
