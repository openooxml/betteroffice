---
'@betteroffice/pptx-react': patch
---

`api.save()` now hands its bytes to `onSave`, as docx and xlsx do. It still never downloads and still skips `onSaveRequest`; before, it only returned the bytes.
