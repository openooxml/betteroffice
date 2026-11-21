---
"@betteroffice/docx": minor
"@betteroffice/docx-react": patch
"@betteroffice/fonts": minor
---

The worker editor opens bundled CJK faces from their skeleton and core and fetches outline chunks for the glyphs each page draws before painting it, so a first CJK frame downloads a fraction of the face; print waits for pending outlines.
