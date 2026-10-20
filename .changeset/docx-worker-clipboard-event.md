---
"@betteroffice/docx-react": patch
---

Worker editor copy and cut write plain text, plus HTML for table selections, to the native clipboard. Pending copies attempt a fresh copy event while user activation is available, then fall back to the async Clipboard API if needed.
