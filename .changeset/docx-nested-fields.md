---
"@betteroffice/docx": patch
---

Documents with deeply nested fields, such as repeated `INCLUDEPICTURE` fields in a header, now open instead of running out of memory or failing to open. In parser output, `fieldTree.code` and `fieldTree.result` are now set only for fields that contain no other field; a field that contains one carries its content in `structuredCode` and `structuredResult` only, and `fieldTree.children` entries no longer repeat code or result.
