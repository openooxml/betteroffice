---
"@betteroffice/docx-react": patch
"@betteroffice/xlsx-react": patch
"@betteroffice/pptx-react": patch
---

On XLSX and PPTX, focusing an input inside a plugin overlay now reports `focused` false, as DOCX does. On every engine, replacing the editor surface while it has focus sends one `focus-change` to false, and an editor opened in an unfocused window starts unfocused.
