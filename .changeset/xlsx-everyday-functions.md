---
"@betteroffice/xlsx": patch
---

Recalculation now computes `RAND`, `FORECAST`, `FORECAST.LINEAR`, `GCD`, `LCM`, `EVEN`, `ODD`, `FACT`, `COMBIN` and `FIXED` instead of returning `#NAME?`, following Excel's argument truncation and error rules.
