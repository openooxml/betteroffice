---
"@betteroffice/docx-react": patch
---

`@betteroffice/docx-react` now re-exports the host-proposal types (`DocxProposalRequest`, `DocxProposalResult`, `DocxProposalStateRequest`, `DocxProposalSnapshot`, `DocxProposalFailure`, and related types), so hosts no longer need to import `@betteroffice/docx/yrs` directly.
