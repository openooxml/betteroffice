---
'@betteroffice/agents': patch
'@betteroffice/docx': patch
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-react': patch
---

Attached agent proposals use the shared engine; standalone behavior is unchanged. Adds opt-in DOCX `editPeer`, `proposeAsync`, `rejectAsync`, `discardAsync`, `keepRevisions`, typed `TrackedRevisionConflictError` and `AcceptDateUnsupportedError`, and XLSX `getProposalStatus` and `onProposalChange`.
