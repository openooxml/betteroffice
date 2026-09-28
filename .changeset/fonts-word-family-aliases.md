---
"@betteroffice/fonts": patch
---

Export `WORD_FAMILY_ALIASES`, the map from Word family names and their normalized spellings to the canonical Word family they stand in for (for example `helvetica` to `arial`); `resolveMetricCompatFamily` then picks the bundled face for that family.
