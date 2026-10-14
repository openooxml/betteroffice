---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Worker-only DOCX editors send keys typed while earlier edits are still pending as one edit instead of one per key, so typing catches up faster on a busy worker. Keys on either side of an IME composition stay separate edits.
