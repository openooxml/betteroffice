---
"@betteroffice/docx": minor
"@betteroffice/rust-crates": minor
---

Round-trip DOCX packages as a byte-stable fixed point that keeps the authored body sectPr, simple fields, drawing names and z-order, foreign markup, unknown attributes, and custom root bindings, modelled through new public fields and enum variants that exhaustive matches and literal constructors must account for.
