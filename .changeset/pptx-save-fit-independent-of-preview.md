---
"@betteroffice/pptx": patch
"@betteroffice/agents": patch
---

Agent presentation exports no longer depend on a prior slide preview: edited autofit text is always shrunk to fit on save, as PowerPoint does. `PresentationHandle.save` accepts `{ fonts, fallbackFonts }` to fit with those faces.
