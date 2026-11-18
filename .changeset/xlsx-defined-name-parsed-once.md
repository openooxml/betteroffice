---
"@betteroffice/rust-crates": patch
"@betteroffice/xlsx": patch
---

A formula evaluated outside a recalculation now parses each defined name's definition once per evaluation context instead of on every use.
