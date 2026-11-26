---
'@betteroffice/docx': patch
---

DOCX files whose XML writes standard namespaces under generated prefixes such as `ns0:` and `ns4:`, as Python's ElementTree saves them, now open with their headers and footers instead of failing with `story "hf:" was not found`.
