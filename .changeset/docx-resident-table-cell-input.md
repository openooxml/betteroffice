---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Typing and deleting in table cells now goes through the resident engine worker in one request, like body paragraphs, instead of laying out the document first.
