---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Keep hyperlinks, fields, and every other unmodeled run markup when a text edit spans several source runs. Surviving text is written back onto the run it came from, typed text extends the run before it, a field whose text changed becomes a plain run so PowerPoint does not overwrite it, and a run whose text is deleted takes its markup with it.
