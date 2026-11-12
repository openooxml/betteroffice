---
'@betteroffice/docx-native': patch
---

`install_panic_hook` now works and the package installs it on load: a Rust panic writes a `[docx-native] panic:` line to stderr, then Rust's default report. The thrown error is unchanged; calling it again does nothing.
