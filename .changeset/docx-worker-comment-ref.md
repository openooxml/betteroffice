---
"@betteroffice/docx-react": patch
"@betteroffice/docx": patch
---

The default worker editor supports comment navigation, replies, resolution and repeated document loads, with superseded navigation returning false. Imported resolved comments can reopen, replies inherit ancestor anchors, and saves await queued comment operations; updates report full projections in `editState.commentChange`.
