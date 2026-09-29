---
"@betteroffice/docx": minor
"@betteroffice/docx-react": minor
---

Add host proposals: `YrsSession.proposeChanges()` records tracked changes under host-proposal ids; `setProposalStates()` previews accept/reject/restore without changing the document, and `getProposals()`/`onProposalChange()` read them. `DocxEditorRef` gains matching methods, and the new `allowHostProposals` prop (off by default) enables them in read-only viewers.
