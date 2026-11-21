---
'@betteroffice/pptx-react': patch
---

Undoing the deletion of a presentation's only slide paints the restored slide again. Before, the worker editor did not repaint it until another edit.
