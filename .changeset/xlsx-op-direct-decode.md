---
'@betteroffice/xlsx': patch
---

Smaller WebAssembly core and faster `applyOps`. Operation JSON whose `type` key is not first decodes more slowly. Malformed operation JSON may report different errors. Apps sending their own operation JSON across package versions should send operations as this package serializes them.
