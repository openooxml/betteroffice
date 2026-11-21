---
'@betteroffice/agents': minor
'@betteroffice/docx': minor
'@betteroffice/docx-react': minor
'@betteroffice/xlsx': minor
'@betteroffice/xlsx-react': minor
---

Attached agent proposals use the shared engine; standalone behavior is unchanged. Adds opt-in DOCX `editPeer`, `proposeAsync`, `rejectAsync`, `discardAsync`, `keepRevisions`, `replaceText.at`, typed `TrackedRevisionConflictError`, `AcceptDateUnsupportedError` and `ProposalDecidedInEditorError`, and XLSX `getProposalStatus`, `onProposalChange` and `isProposalStale`.
