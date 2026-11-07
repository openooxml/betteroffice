---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

Browser `renderPng` without `fonts` paints with the bundled faces native uses, so canvas and native PNGs match; a face that cannot load rejects with `font-unavailable`. Native chart text in a family list (`Calibri, sans-serif`) now resolves its first family.
