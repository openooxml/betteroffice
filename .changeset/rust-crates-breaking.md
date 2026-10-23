---
"@betteroffice/rust-crates": minor
---

**Breaking:** Public structs and enums gain fields and variants, including `betteroffice-xlsx` `CalculationResult.budget_exhausted` and types the DOCX and PPTX crates re-export, such as `DisplayPage` and `SlideSnapshot`. Update struct literals and exhaustive matches. See https://docs.betteroffice.dev/docs/migrating-to-0-5.
