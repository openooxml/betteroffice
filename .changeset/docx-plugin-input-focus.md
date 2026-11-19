---
'@betteroffice/docx-react': minor
---

Plugins can read `snapshot.inputFocused` and return focus to the document with `context.focusDocument()`, which keeps the selection. Both members are required, so hand-built snapshots and contexts (e.g. test doubles) must add them.
