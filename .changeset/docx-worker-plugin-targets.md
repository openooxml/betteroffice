---
'@betteroffice/docx': minor
'@betteroffice/docx-react': minor
---

Worker-only editors connect plugin comments and navigation: `insertComment` resolves to the new comment's ID, also in suggesting mode and right after a save, and `scrollToParagraph` moves the selection only with `focus`. Saves keep comment bodies, replies and resolved threads.
