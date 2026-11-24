---
"@betteroffice/agents": patch
---

`xlsx_sort` now moves formulas with their rows, so a formula that refers to its own row still does after the sort. Sorts that could leave moved formulas with stale results or change ambiguous references are refused.
