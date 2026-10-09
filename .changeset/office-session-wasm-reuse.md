---
"@betteroffice/pptx": patch
"@betteroffice/xlsx": patch
---

Built-in worker sessions for presentations and workbooks start compiling the engine as soon as the worker starts, and later built-in sessions in the same page reuse the compiled module, so opening another document is faster.
