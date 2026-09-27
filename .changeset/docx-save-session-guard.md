---
"@betteroffice/docx-react": patch
---

Save no longer fails with "The document changed while saving" when the editor lays out again during the save; it still aborts when another document is loaded meanwhile. A failed input operation now fails only the flush, save or command that was waiting for it, so Save, print and toolbar commands keep working for the rest of the session.
