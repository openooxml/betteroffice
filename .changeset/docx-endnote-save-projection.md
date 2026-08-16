---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Edits made in an endnote now reach the saved file. Endnote stories were seeded into the collaborative document but never projected back on save, so every endnote edit was silently replaced by the imported text; footnotes and endnotes now project through one path, and typing in a note or header — including a table cell inside one — marks its root story for the changed-stories-only save.
