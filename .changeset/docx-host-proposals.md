---
"@betteroffice/docx": minor
"@betteroffice/docx-react": minor
---

Add host proposals: `YrsSession.proposeChanges()` records a tracked-change batch under host-proposal ids, outside undo history. `setProposalStates()` previews accept/reject/restore via the new `revisionPreview` entry without changing the document; `getProposals()`/`onProposalChange()` read them, and `DocxEditorRef` with the new `allowHostProposals` prop (off by default) expose them to read-only viewers.
