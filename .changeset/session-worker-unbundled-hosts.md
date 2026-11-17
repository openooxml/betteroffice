---
"@betteroffice/xlsx": patch
"@betteroffice/pptx": patch
"@betteroffice/docx": patch
"@betteroffice/xlsx-native": patch
"@betteroffice/pptx-native": patch
"@betteroffice/docx-native": patch
---

Worker files and main-thread entries import no private `#` specifiers, so worker sessions open from a CDN or with an import map, without a bundler. On Node.js, a worker thread that fails to start now fails the session instead of throwing synchronously.
