---
'@betteroffice/pptx': minor
'@betteroffice/pptx-react': minor
'@betteroffice/fonts': minor
---

Editors given `fonts` now load bundled faces for the other families a deck uses, including light members such as Calibri Light, so their text keeps its metrics. Supplied families keep the app's faces. Pass `defaultFonts={false}` to fetch no default fonts.
