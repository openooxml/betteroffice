---
"@betteroffice/xlsx": patch
---

Dashboards that sum large data sheets with `SUMPRODUCT((range=key)*...*values)` now calculate every cell instead of stopping with #NUM! partway through.
