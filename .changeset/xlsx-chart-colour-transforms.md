---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/python-xlsx': patch
'@betteroffice/rust-crates': patch
---

XLSX chart colours apply their `lumMod`, `lumOff`, `tint`, `shade` and other colour transforms, so lighter or darker series shades and grey axis text draw as Excel shows them. `sysClr` and `prstClr` colours resolve, and translucent colours keep their alpha.
