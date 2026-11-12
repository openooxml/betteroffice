---
'@betteroffice/xlsx-native': minor
---

Adds `install_panic_hook`, matching the WebAssembly engine. The package installs it on load, so on macOS and Linux a Rust panic writes one `[xlsx-native] panic:` line to stderr, then Rust's default report. The thrown error is unchanged; calling it again does nothing.
