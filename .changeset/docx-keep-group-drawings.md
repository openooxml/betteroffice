---
"@betteroffice/docx": patch
"@betteroffice/docx-native": patch
---

Keeps drawing groups and canvases on save in headers, footers and rewritten paragraphs. A group that matches its source apart from its frame and text-box text is written back from the source with those changes; any other edit writes it from the model as before.
