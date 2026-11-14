---
"@betteroffice/xlsx": patch
---

Smaller WebAssembly cores. Behaviour change: after `initWasm(input, { png: false })`, the standalone `renderXlsxMarkdown` now rejects, as PNG export already does there. Sessions, `exportMarkdown` and the default init are unchanged.
