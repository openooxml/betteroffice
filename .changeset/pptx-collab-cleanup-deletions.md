---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Collaborating deck replicas no longer drift apart in text formatting. Deletions a replica makes while cleaning up formatting during a remote update now arrive as their own `'local'` update, which providers forward.
