---
"@betteroffice/xlsx": patch
"@betteroffice/pptx": patch
"@betteroffice/docx": patch
"@betteroffice/xlsx-native": patch
"@betteroffice/pptx-native": patch
"@betteroffice/docx-native": patch
---

Worker entries no longer import the module that constructs their own worker, so Next.js (Turbopack) production builds of apps that use `./worker` finish instead of hanging.
