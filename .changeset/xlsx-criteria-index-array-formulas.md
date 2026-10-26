---
"@betteroffice/xlsx": patch
---

Dashboards that sum large data sheets with `SUMPRODUCT((range=key)*...*values)` inside array (Ctrl+Shift+Enter) formulas now calculate every cell instead of stopping with #NUM! partway through.
