---
"@betteroffice/docx-react": patch
---

The default worker-only editor restores checkpoints and replays edits once after a crash or out-of-memory failure outside collaborative sessions. Collaborative failures, second failures and refused replays reach `onError`; replacement-worker OOM errors retain their classification and memory diagnostics.
