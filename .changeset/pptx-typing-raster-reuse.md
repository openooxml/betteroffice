---
'@betteroffice/pptx-react': patch
---

Typing in a PowerPoint text box no longer allocates and repaints a full slide canvas per keystroke: the editor reuses the slide raster and repaints only the lines that changed.
