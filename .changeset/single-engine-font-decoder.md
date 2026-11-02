---
"@betteroffice/fonts": patch
"@betteroffice/docx-react": patch
"@betteroffice/pptx-react": patch
"@betteroffice/xlsx-react": patch
"@betteroffice/demo": patch
---

Editors decode packed fonts with their engine by default instead of downloading a separate decoder. Standalone font providers can opt into a decoder with `runtime.decoder`.
