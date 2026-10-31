---
"@betteroffice/docx": patch
"@betteroffice/agents": patch
---

Preserves SmartArt during agent and editor saves, including resized diagrams and large source payloads. `Shape.drawingXml` or `Shape.drawingRunXml` stores the source once; `Shape.drawingModel` tracks frame edits. Enabled by default.
