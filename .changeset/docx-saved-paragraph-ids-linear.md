---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Saving a document records paragraph ids in linear time instead of comparing every saved id with every live paragraph. At 5 MiB this step drops from about 2.7 s to 63 ms on native and 4.9 s to 0.2 s on wasm; output is unchanged.
