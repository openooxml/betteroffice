---
'@betteroffice/python-docx': minor
---

`Document.save_report()` takes the same keywords as `save()` and returns a `SaveReport`: the bytes `save()` writes, the paragraphs kept as source bytes in `kept_unchanged`, and the structure a part written whole dropped in `structure_dropped`.
