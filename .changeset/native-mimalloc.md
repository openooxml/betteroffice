---
'@betteroffice/docx-native': minor
'@betteroffice/pptx-native': minor
---

The docx and pptx native bindings use mimalloc as their memory allocator, through the default-on `mimalloc` cargo feature. DOCX save, edit and layout get faster and PPTX open and export faster, with a higher peak memory on some documents.
