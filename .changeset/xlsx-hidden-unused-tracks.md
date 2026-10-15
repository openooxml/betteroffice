---
"@betteroffice/xlsx": patch
---

Renders sheets that hide their unused columns or rows. A frame now ends at the last column and row with extent, so these sheets no longer fail with `DisplayTooLargeError` at the first viewport.
