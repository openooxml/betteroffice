---
"@betteroffice/docx": minor
"@betteroffice/docx-react": minor
---

DOCX plugins can return a `DocxReservationSet` from the opt-in `reservations` hook to lay out inline boxes, blocks and hidden ranges, then draw in them through `getAnchorGeometry({ kind: 'reservation', id })` once `DocxPluginLayout.reservationsVersion` is the set's version. Saves and paged exports leave them out.
