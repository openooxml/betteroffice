---
"@betteroffice/docx": patch
"@betteroffice/docx-native": patch
---

Keeps drawing groups and drawing canvases on save in headers, footers and rewritten paragraphs. When a group matches its source drawing apart from its frame and text-box text, the source is written back with those changes applied; any other edit writes the group from the model as before.
