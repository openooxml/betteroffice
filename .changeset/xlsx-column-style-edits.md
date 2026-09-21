---
'@betteroffice/xlsx': patch
'@betteroffice/rust-crates': minor
'@betteroffice/python-xlsx': patch
---

Keep a column style attached to its columns through an edit, and make it part of what identifies a workbook. Inserting or deleting columns now moves the `<col>` runs with the cells: a run the insert splits widens over the new columns, a run to its right shifts, and a delete clips a run to the columns that survive and drops one it consumes whole. Undoing that delete restores a clipped run exactly; a run the delete consumed whole does not come back, because no operation carries column styles yet. The collaboration fingerprint hashes column styles at the current schema version, so two peers whose workbooks differ only there no longer share a session and render differently; a workbook that declares no column style keeps the fingerprint every earlier release computed, and one that declares some carries that pre-change fingerprint in its accepted set, so saved snapshots still attach. A worksheet may declare at most 65,536 style-bearing `<col>` runs, refused as `TooManyColumnStyles` in the shape of the existing cell, hyperlink and chart caps.
