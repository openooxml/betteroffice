---
'@betteroffice/pptx': minor
'@betteroffice/pptx-react': patch
---

Text composed with an input method is now typed into the slide as it is composed, in the slide's font. The command protocol gains `compositionUpdate`; a composition is one undo step, and cancelling one typed over a selection gives the selection back.
