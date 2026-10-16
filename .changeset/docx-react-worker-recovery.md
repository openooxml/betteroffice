---
"@betteroffice/docx-react": patch
---

The default worker-only editor replaces a crashed or out-of-memory worker once, reopening a `documentBuffer` at its last checkpoint and replaying later edits. A second failure reaches `onError`.
