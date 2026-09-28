---
"@betteroffice/docx": minor
"@betteroffice/docx-react": minor
---

TypeScript changes to check when upgrading from 0.2: parsed chart `ChartSeries.values`, `xValues` and `bubbleSizes` are `(number | null)[]`, with `null` for a point a sparse cache leaves empty; `YrsTableReceipt.changedStoryIds` is a required field, so receipt literals and mocks need it; `YrsResidentWorkerSnapshot.fonts` holds `YrsResidentFontRegistration` entries instead of raw bytes. In suggesting mode the editor now refuses edits it cannot record as tracked changes (inserting tables, page and section breaks, image layout and properties, page setup, watermarks, and table actions other than inserting and deleting rows), where 0.2 inserted tables as tracked rows and applied the rest untracked. Cmd/Ctrl+S now runs the editor's Save, including in read-only mode, and downloads the file unless `downloadOnSave` is `false`; hosts that handle the shortcut themselves take it over with `onSaveRequest` or by calling `preventDefault()` in a capture-phase listener.
