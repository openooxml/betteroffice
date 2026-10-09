---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Preserves opaque drawings, tracked changes and hyperlinks through editor saves by default, and recovers legacy chart placements. Adds opt-in `repackDocxWithWarnings` with `warnings`; seeding rejects opaque payloads over 8 MiB with `OpaqueSeedBudgetError`.
