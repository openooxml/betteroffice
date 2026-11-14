---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Opening a DOCX keeps equal embed keys and values in one shared allocation: retained memory after open falls 7.8 to 12.0% on the two largest public test documents and 29.2 to 30.5% on a 5 MiB synthetic document, native and wasm. Saved bytes are unchanged.
