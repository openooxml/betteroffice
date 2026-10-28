---
"@betteroffice/pptx": patch
"@betteroffice/pptx-react": patch
---

Lower PowerPoint editor memory on large decks: crash recovery keeps one copy of the opening state and the edits since then, instead of re-encoding and copying the whole presentation at every automatic checkpoint.
