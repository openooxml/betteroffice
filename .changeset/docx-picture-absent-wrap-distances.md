---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Saving no longer adds zero wrap distances or a zero effect extent to a picture that had none, so inline pictures keep their position in LibreOffice after a save.
