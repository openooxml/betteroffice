---
"@betteroffice/pptx": patch
---

Editing text no longer drops authored formatting: unchanged fonts keep `a:latin` panose, pitchFamily and charset, unchanged colours keep theme tokens and transforms instead of becoming RGB, styled breaks keep their own `rPr`, and changed font families drop old face metrics.
