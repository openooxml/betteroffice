---
"@betteroffice/xlsx": patch
---

Printed sheets now size rows and default columns the way Excel's print does. Rows without a height take the font's printed height, stored heights land on whole points, and unsized columns are as wide as the base width implies, so content no longer drifts up.
