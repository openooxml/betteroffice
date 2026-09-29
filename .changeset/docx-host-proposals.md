---
"@betteroffice/docx": minor
"@betteroffice/docx-react": minor
---

Add host proposals: `YrsSession.proposeChanges()` records a tracked-change batch outside undo history, and `setProposalStates()` previews accept/reject/restore via the new `revisionPreview` entry without changing the document. `getProposals()` and `YrsSession.onProposalChange()` read them; `DocxEditorRef`'s matching methods reach read-only viewers via the new `allowHostProposals` prop, off by default.
