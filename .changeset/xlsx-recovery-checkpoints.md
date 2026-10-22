---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
---

Worker-only workbook sessions no longer fail after 256 edits or selection moves without a recovery checkpoint. The session now checkpoints its recovery log itself, including headless sessions, and a checkpoint that fails reports a typed session failure.
