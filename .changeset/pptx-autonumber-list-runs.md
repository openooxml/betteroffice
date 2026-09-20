---
'@betteroffice/pptx': patch
---

Number `a:buAutoNum` paragraphs as one list when they repeat the same `startAt`: PowerPoint writes a list's start on every one of its paragraphs, so a four-item list marked `startAt="4"` now draws 4, 5, 6, 7 and one marked `startAt="1"` draws a), b), c). A paragraph that declares a different start still opens a new list.
