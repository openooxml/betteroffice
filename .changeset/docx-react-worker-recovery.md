---
"@betteroffice/docx-react": patch
---

The default worker-only editor replaces a crashed or out-of-memory worker once, restoring its checkpoint and replaying later edits. A second failure or a refused replay reaches `onError`; replacement-worker OOM errors retain their classification and memory diagnostics.
