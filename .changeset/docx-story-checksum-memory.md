---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
---

Story checksums use constant memory, so long documents no longer exhaust the WebAssembly engine's memory and stop the editor. Checksum values are unchanged.
