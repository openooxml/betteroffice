---
"@betteroffice/rust-crates": minor
---

**Breaking:** `betteroffice-xlsx` `CalculationResult` adds the public `budget_exhausted` field, so struct literals must set it. The engine crates add public fields, enum variants and parameters; update struct literals and exhaustive matches. See https://docs.betteroffice.dev/docs/migrating-to-0-5.
