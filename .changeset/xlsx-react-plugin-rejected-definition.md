---
'@betteroffice/xlsx-react': patch
---

When a rejected plugin definition shares its id with an accepted one, edit gating and row decorations now use the accepted plugin's hooks instead of the rejected one's.
