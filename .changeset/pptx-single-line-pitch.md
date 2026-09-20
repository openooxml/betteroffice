---
'@betteroffice/pptx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-pptx': patch
---

Measure a single-spaced PowerPoint line as 1.2 em of the line's largest font, the pitch PowerPoint 16.113 uses for every face, so percentage line spacing no longer inherits the substituted face's own ascent, descent and line gap and multi-line bodies stop drifting away from PowerPoint down the shape. Super- and subscript ink still pushes the line box out past that pitch.
