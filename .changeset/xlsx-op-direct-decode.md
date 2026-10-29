---
'@betteroffice/xlsx': patch
---

Smaller WebAssembly core and faster `applyOps`. Operations whose `type` key is not first take a slower fallback, so put `type` first. Malformed operation JSON may report different errors. Apps sending their own operation JSON across package versions should send operations as this package serializes them.
