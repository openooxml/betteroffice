---
'@betteroffice/xlsx-native': minor
---

Adds `install_panic_hook`, matching the WebAssembly engine. The package installs it on load, so a Rust panic writes an `[xlsx-native] panic:` line to stderr ahead of Rust's default report. The thrown error is unchanged; calling it again does nothing.
