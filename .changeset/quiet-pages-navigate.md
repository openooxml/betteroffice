---
"@betteroffice/docx-react": patch
---

Deprecates `scrollToPosition` in favor of `scrollToParagraph` and `highlightRange` in favor of `readPositionAtPoint`, with accurate worker viewer warnings. Preserves documented synchronous viewer ref results and reports refused queued `proposeChange` calls through `onError` by default.
