---
"@betteroffice/docx": patch
---

`withdrawProposals` now refuses with `tracked-revision-conflict` when a proposal's tracked change also holds edits made outside the proposals, instead of removing those edits.
