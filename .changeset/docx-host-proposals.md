---
"@betteroffice/docx": minor
"@betteroffice/docx-react": minor
---

Add host proposals. `YrsSession.proposeChanges()` records a round of tracked changes under host proposal ids, resolved by paragraph anchor and search with a one-based or `'all'` non-overlapping occurrence, as one batch outside undo history or a typed refusal with nothing changed. `setProposalStates()` previews accepted, rejected and restored proposals through the new `revisionPreview` render environment entry without changing the document, its version or undo history, and `getProposals()`/`onProposalChange()` read and observe the registry. `DocxEditorRef` gains the matching async methods, and the `allowHostProposals` prop admits them in a read-only viewer.
