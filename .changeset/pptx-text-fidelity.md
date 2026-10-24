---
"@betteroffice/pptx": patch
---

Fixes unnecessary inherited autofit overrides during text edits while saving changed fits for overflowing text. `setStoryParagraphs` now accepts line feeds within runs as soft line breaks.
