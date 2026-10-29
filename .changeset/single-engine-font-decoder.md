---
"@betteroffice/fonts": patch
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/pptx": patch
"@betteroffice/pptx-react": patch
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
"@betteroffice/vsdx": patch
"@betteroffice/demo": patch
---

Editors decode packed fonts with their engine by default instead of downloading a separate decoder. Standalone font providers can opt into a decoder with `runtime.decoder`.
