---
'@betteroffice/pptx': patch
'@betteroffice/rust-crates': patch
---

Inserting or duplicating many slides in a presentation opened with `Presentation::open` no longer replaces an existing slide when a new slide's id matches one from the file; each call adds one slide.
