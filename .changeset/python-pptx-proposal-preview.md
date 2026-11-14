---
'@betteroffice/python-pptx': minor
---

`Presentation.proposal_slide_strip` and `layout_proposal_diff_slide_by_id` match the TypeScript proposal preview API, and `render_proposal` resolves slide IDs and indices in the proposed deck, so inserted slides render. `Proposal.refs` and `ProposalChange.kind`, `before_index` and `after_index` carry the structural proposal fields.
