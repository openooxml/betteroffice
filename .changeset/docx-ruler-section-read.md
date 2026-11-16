---
'@betteroffice/docx': minor
'@betteroffice/docx-react': patch
---

The `sectionProperties` read takes an optional `paragraphs` list that limits the tab stops it returns to those paragraphs. The ruler uses it, so typing in long documents no longer waits for every paragraph to be re-read after each keystroke.
