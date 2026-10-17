---
"@betteroffice/docx": patch
---

Documents with deeply nested fields (e.g. repeated header `INCLUDEPICTURE` fields) now open instead of failing or running out of memory. Parser output sets `fieldTree.code` and `fieldTree.result` only on fields that contain no other field; outer fields keep their content in `structuredCode` and `structuredResult`.
