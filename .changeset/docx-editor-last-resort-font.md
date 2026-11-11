---
'@betteroffice/docx': minor
---

`TextMeasureFontRegistry` accepts `standInTail`, so explicit stand-ins always append the last-resort face. With no font setup, the editor draws fonts it lacks, such as Aptos, with a bundled stand-in, so lines no longer break mid-word. The console warns once per font until the fonts are reconfigured.
