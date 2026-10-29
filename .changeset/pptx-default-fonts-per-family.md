---
'@betteroffice/pptx': minor
'@betteroffice/pptx-react': minor
---

Editors given `fonts` now load bundled faces for the other families a deck uses, so Times New Roman, Calibri or CJK text keeps its metrics. Supplied families keep the app's faces. Pass `defaultFonts={false}` to fetch no default fonts.
