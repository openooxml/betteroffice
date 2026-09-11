---
"@betteroffice/rust-crates": minor
---

Intern styles through hash indexes instead of linear scans. `Stylesheet` gains private interning caches, which prevent exhaustive struct-literal construction downstream; the caches are internal state only and excluded from serialization and equality.
