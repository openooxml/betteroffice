---
"@betteroffice/xlsx": patch
---

Dashboards that sum large data sheets with `SUMPRODUCT((range=key)*...*values)` now calculate every cell instead of stopping with #NUM! partway through: repeated lookups over the same constant columns reuse one index per recalculation, so they cost far less work.
