---
'@betteroffice/pptx-react': patch
---

Keep the shape selection outline a constant-width rectangle while resizing. The worker overlay now positions the outline from the preview bounds instead of applying a non-uniform CSS scale to the original box, so the 2px border no longer stretches on one axis during a drag.
