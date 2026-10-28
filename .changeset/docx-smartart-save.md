---
"@betteroffice/docx": patch
"@betteroffice/agents": patch
"@betteroffice/rust-crates": patch
---

Preserves inline and anchored SmartArt when editing text in the same paragraph through agent or editor saves. Source markup is retained by default through `Shape.drawingXml` and `Shape.drawingRunXml`, keeping diagram relationships and metadata intact.
