---
'@betteroffice/rust-crates': patch
'@betteroffice/xlsx': patch
---

Time formats with many fractional-second digits, such as `ss.` followed by dozens of zeros, no longer crash or show wrong digits. Seconds round to milliseconds as in Excel, and further digits show as zeros.
