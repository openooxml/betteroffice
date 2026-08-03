---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
"@betteroffice/rust-crates": patch
---

A workbook snapshot persisted by the previous release opens again. A replica bootstraps the workbook it opened into a document whose client ID is the head of the base fingerprint, and that fingerprint hashes the collaboration schema version — so raising the version for charts gave every workbook a different bootstrap identity. A snapshot an earlier release wrote no longer deduplicated against the one this build seeds: the two bases doubled up, one was tombstoned by client ID, and restoring reported that the shared workbook structure had changed or silently handed back the pristine file.

A replica that has not been edited yet now takes a whole snapshot as its state and upgrades it in place, rather than merging it against a bootstrap it can never agree with. Where the two bootstraps do agree the snapshot and the merge describe the same document, so this is never the worse answer. The upgraded state, not the snapshot, is what peers are told about: the upgrade writes new structs, and an incremental update that later builds on them would sit unintegrated forever on a peer that never received them. What the frozen structure describes — sheet order and names, merges, freeze panes, hyperlinks and charts — must still match, disregarding the shared-type identities a replaced bootstrap changes by construction. A snapshot that fails that, or a whole document this build cannot read, is now an error rather than a silent no-op.

A charted workbook pairs with a pre-chart snapshot again. Such a snapshot carries no chart state to disagree about — charts come from the file the replica opened, keyed to the sheet they were parsed from — so refusing the pairing only made charted workbooks the ones that could never be restored.

Workbooks with a hidden row or column restore too. This release began modelling those as a zero dimension where earlier ones recorded nothing at all, and both dimension maps are fingerprinted, so such a workbook could not be recognised as the base its own snapshot had started from. The dimensions an earlier release would have stored are now read from the source sheet alongside the current ones and accepted as a legacy fingerprint, and restoring puts the hidden dimensions back rather than letting the row silently unhide.

Each feature is now pinned to the schema that introduced it rather than to whichever schema is current, so the next version bump cannot reclassify the newest schema as predating charts and manufacture this same failure again.
