---
"@betteroffice/agents": minor
---

`pptx_edit` `move_slide` takes `after` or `before` a slide ID as an alternative to `index`, so a slide can be placed next to another without position arithmetic. Saved decks keep every slide's `p:sldId` across reorders, and inserted slides take IDs above all existing ones.
