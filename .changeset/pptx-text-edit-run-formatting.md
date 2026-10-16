---
"@betteroffice/pptx": patch
---

Editing run text no longer drops authored formatting. An unchanged font keeps its `a:latin` panose, pitchFamily and charset, and an unchanged colour keeps its theme colour and transforms instead of becoming RGB. A changed font family drops the old face metrics.
