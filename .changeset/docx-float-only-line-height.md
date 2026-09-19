---
'@betteroffice/docx': patch
---

Size a line that carries only an anchored float from the paragraph mark rather than a synthetic 0.8/0.2 em box. Word measures such a line exactly as it measures an empty paragraph, so the old fallback made every float-only paragraph about a point short and let one paragraph too many onto the page.
