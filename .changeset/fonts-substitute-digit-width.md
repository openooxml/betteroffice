---
"@betteroffice/fonts": patch
---

Adds optional `bold` and `italic` arguments to `maxDigitWidth` and `resolveDigitAdvance` so Verdana column widths use the requested style's digit metrics. Both default to `false`; unrecorded styles continue measuring the drawn face.
