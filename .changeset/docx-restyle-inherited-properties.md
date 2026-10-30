---
'@betteroffice/docx': patch
'@betteroffice/agents': patch
---

`set_paragraph_style` keeps borders, spacing and other properties inherited from the new style when saving an unnumbered paragraph, so later style changes continue to apply.
