---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

A document whose chart part cannot be read, or is absent from the package, now opens with that chart missing instead of the whole file being refused, and the drawing that referenced it is carried as opaque content — a full save drops it, as it already drops every chart, rather than writing it back as a picture pointing at whatever `rId1` happens to be. Chart parts read against one shared event allowance, so a document cannot spend more on charts than one part was allowed before, and the part the parser declined is still written back untouched on save.
