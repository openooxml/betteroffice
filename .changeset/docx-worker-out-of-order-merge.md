---
"@betteroffice/docx": patch
---

Worker-hosted DOCX collaboration no longer fails when a remote update arrives before the update it depends on: the replica keeps its last view until the missing update arrives, then converges.
