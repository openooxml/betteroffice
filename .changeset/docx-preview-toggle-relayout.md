---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Reduces relayout work for proposal previews changed through `setProposalStates`. `ResidentEngineWorkerClient.sync` takes an opt-in `supersedable` option that lets a newer queued sync replace it.
