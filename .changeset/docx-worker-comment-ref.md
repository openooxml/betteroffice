---
"@betteroffice/docx-react": patch
"@betteroffice/docx": patch
---

The default worker editor supports comment navigation, replies, resolution and repeated document loads. Navigation returns false when superseded. Imported resolved threads can reopen, replies inherit ancestor anchors, and saves include queued comment changes. Comment changes report the full projection in `editState.commentChange`.
