---
"@betteroffice/docx": patch
"@betteroffice/agents": patch
"@betteroffice/rust-crates": minor
---

Preserves SmartArt during agent and editor saves, including resized diagrams and large source payloads. `Shape.drawingXml` or `Shape.drawingRunXml` stores the source once; `Shape.drawingModel` tracks frame edits. Enabled by default. Rust `Shape` literals must set `drawing_xml`, `drawing_run_xml` and `drawing_model`.
