---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

The editor no longer loads, or reports errors for, East Asian or complex-script fonts a document names without having text in that script. `YrsDocxHost` gains `unusedScriptFonts`, which lists them.
