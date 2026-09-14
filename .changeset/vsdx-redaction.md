---
"@betteroffice/rust-crates": patch
---

Redact Visio packages instead of refusing them. `ooxml-redact` now accepts `.vsdx` and `.vstx`,
removing shape text and the Value, Prompt and Label cells of Property and User sections while
preserving the structure and geometry the package needs to stay valid. Macro-enabled files stay
refused, and a package whose ShapeSheet nesting cannot be resolved is refused rather than partially
redacted.
