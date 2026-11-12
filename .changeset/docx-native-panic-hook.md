---
'@betteroffice/docx-native': patch
---

`install_panic_hook` now works and the package installs it on load: on macOS and Linux a Rust panic writes one `[docx-native] panic:` line to stderr, then Rust's default report. The thrown error is unchanged; calling it again does nothing.
