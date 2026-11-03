---
'@betteroffice/agents': patch
'@betteroffice/docx': patch
'@betteroffice/docx-react': patch
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-react': patch
---

Attached agent proposals use the shared engine; standalone behavior is unchanged. Adds opt-in DOCX `editPeer`, `proposeAsync`, `rejectAsync`, `discardAsync`, `keepRevisions`, `replaceText.at`, typed `TrackedRevisionConflictError` and `AcceptDateUnsupportedError`, and XLSX `getProposalStatus`, `onProposalChange` and `isProposalStale`.
