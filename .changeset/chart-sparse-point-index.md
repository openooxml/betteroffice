---
"@betteroffice/docx": patch
"@betteroffice/pptx": patch
"@betteroffice/xlsx": patch
"@betteroffice/python-docx": patch
"@betteroffice/python-pptx": patch
"@betteroffice/python-xlsx": patch
"@betteroffice/rust-crates": patch
---

Charts whose cached data names a very large point index no longer allocate a slot for every index up to it, which let a small chart use hundreds of megabytes and run WebAssembly out of memory. A cache now spans no more than its `c:ptCount`, points past it are ignored, and empty slots between points count against the chart's point limit; missing points still draw as gaps.
