---
"@betteroffice/docx": patch
---

Documents with deeply nested fields, such as repeated header `INCLUDEPICTURE` fields, now open instead of failing or running out of memory. `fieldTree.code` and `fieldTree.result` are set only on innermost fields; outer fields keep their content in `structuredCode` and `structuredResult`.
