---
'@betteroffice/agents': patch
'@betteroffice/docx': patch
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-react': patch
---

Agent proposals now use engine proposals and reflect editor decisions. Adds opt-in DOCX `editPeer` with `proposeAsync`, `rejectAsync`, and `discardAsync`, XLSX `getProposalStatus` and `onProposalChange`, and opt-in DOCX `keepRevisions` to retain tracked revisions.
