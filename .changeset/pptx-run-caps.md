---
'@betteroffice/pptx': patch
'@betteroffice/python-pptx': patch
'@betteroffice/rust-crates': patch
---

Support `a:rPr/@cap`, so a run that asks for all caps or small caps is drawn that way. `all` uppercases the run for drawing and `small` also draws the lowercase stretches at Word's 0.8× small-cap size; `none` turns off an inherited setting. The value cascades from a slide master's and layout's `a:defRPr` the way the other run properties do, and the uppercasing itself is now shared with DOCX rather than reimplemented. Casing is a display property: the stored run text keeps the author's casing, so the editor's story, the caret offsets it reports and the saved package are all unchanged — an untouched deck still saves byte-identically. Measured against the PPTArena corpus tail, where `cap="all"` reaches titles in `pptarena-051`, `pptarena-052`, `pptarena-054` and body runs in `pptarena-034`.
