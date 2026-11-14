---
'@betteroffice/rust-crates': patch
'@betteroffice/xlsx': patch
---

Conditional format rules and groups holding a nested element of the same name are now read whole, ending only at their own closing tag, so a later save no longer writes them truncated. Content after the nested closing tag is now read.
