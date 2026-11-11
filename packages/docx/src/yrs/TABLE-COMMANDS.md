# Editor commands and unavailable tables

A table that crossing undos left without its `tblPr`, `grid` or `rows` is unavailable: the engine
refuses every operation on it with the coded `table-unavailable` error, and it does so before it
writes anything (`read_table` is the first read in each table operation). `applyDocxCommand` turns
that error into a clean `table-unavailable` refusal, which is sound only while no command writes
the document before its first table operation. This table lists every command kind of
`DocxCommandPayloads`; `tableCommands.test.ts` fails when a kind has no row or a `table-op` row
writes first.

- `none`: reaches no table operation that can refuse.
- `table-op`: reaches one; `mutates before` must be `no`.
- `reads-payload`: reads a table payload in TS; an incomplete payload is rejected as shown.

| kind | reaches | engine call | mutates before | incomplete payload |
|---|---|---|---|---|
| applyRemoteUpdate | none | engine.applyUpdate residentCollaboration.ts:114; afterRemoteMerge to repairCommandSelection Actions:1448 (setCellSelection :884 only on projection-live cells) | n/a | a merge can make a table unavailable; repair clears the cell selection (:841) and re-places the caret outside |
| applyRemoteUpdateChunks | none | engine.applyUpdateChunks residentCollaboration.ts:113; same repair | n/a | as applyRemoteUpdate |
| collabPipe | none | transport into the same merge path (residentCollaboration.ts:105) | n/a | as applyRemoteUpdate |
| applyEdits | none | proposalEngine.applyEdits via residentCommandBatches.ts:12; batch steps are text/paragraph/section only (batch.rs) | n/a | DeleteText/DeleteParagraphs remove a table embed whole; no payload read |
| setRenderEnvironment | none | none | n/a | none |
| setFontEnvironment | none | none | n/a | none |
| setContentControlValue | none | engine.setControlValue Actions:974 | n/a | a control inside an unavailable table does not resolve (stale-control) |
| addRepeatingItem | none | engine.repeatControlAt Actions:976 (Rust repeat_control_at) | n/a | Rust repeat_control_at refuses with table-unavailable before writing when the item holds an incomplete table |
| removeRepeatingItem | none | engine.repeatControlAt(remove) Actions:976 | n/a | removes the item and any table in it; works on unavailable tables |
| openPart | none | none | n/a | none |
| ensurePart | none | setPartDocument Actions:1006 (part metadata only) | n/a | none |
| removePart | none | setPartDocument Actions:1014 | n/a | none |
| setHyperlink | none | applyPropertyCommand residentPropertyCommands.ts:150-158 | n/a | none |
| setObjectProperties | none | setImageGeometryAt residentPropertyCommands.ts:174 (image only) | n/a | none |
| setSectionProperties | none | setPropertyDocument residentPropertyCommands.ts:183 | n/a | none |
| setNoteProperties | none | setPropertyDocument residentPropertyCommands.ts:183 | n/a | none |
| setWatermark | none | setPropertyDocument residentPropertyCommands.ts:190 | n/a | none |
| setSectionMargins | none | engine.setSectionMargins Actions:1028 | n/a | none |
| setTabStops | none | setParagraphAttrs Actions:1248 | n/a | none |
| removeTabStop | none | setParagraphAttrs Actions:1258 | n/a | none |
| formatSelection | none | engine.formatRange Actions:1371 | n/a | formattingRangeContainsContainer (Actions:51) reads segment kinds only |
| toggleStrike | none | engine.formatRange Actions:1364 | n/a | none |
| toggleList | none | setParagraphAttrs Actions:1321 | n/a | none |
| changeIndent | none | setParagraphAttrs Actions:1321 | n/a | none |
| insertImage | none | deleteRange Actions:1326, insertImage :1332 | n/a | none |
| insertTable | none | setCellSelection Actions:1341 on the table just created by insertTable :1336 | n/a | its only table call selects the table insertTable just created, complete by construction |
| insertTableRow | table-op | setCellSelection residentCommandTables.ts:223, then insertRow :224 | no | preflight resolveLiveCell (Tables.ts:38) refuses invalid-cell-target for an unavailable table |
| insertTableColumn | table-op | setCellSelection Tables.ts:228, then insertColumn :229 | no | as insertTableRow |
| deleteTableRow | table-op | deleteRow Tables.ts:235; afterwards syncCellSelectionWithCaret :179 | no | as insertTableRow |
| deleteTableColumn | table-op | deleteColumn Tables.ts:236 | no | as insertTableRow |
| deleteTable | none | deleteTable Tables.ts:246 (Rust removes an unavailable table instead of refusing) | n/a | preflight needs a caret in a live cell, so this command can never target an unavailable table; deleting one needs a range delete |
| mergeTableCells | table-op | mergeCells Tables.ts:250 | no | the cell resolves through the projection, which gives an unavailable table no cells; liveAnchors rejects an incomplete payload |
| splitTableCell | table-op | splitCell Tables.ts:257 (liveAnchor :256 is read-only) | no | the cell resolves through the projection, which gives an unavailable table no cells; liveAnchors rejects an incomplete payload |
| setTableCellShading | table-op | setCellShading Tables.ts:262 | no | as insertTableRow |
| setTableBorders | table-op | setCellBorders Tables.ts:265 | no | as insertTableRow |
| setTableProperties | table-op | setTableProperties Tables.ts:268 | no | as insertTableRow |
| pasteFragment | none | applyClipboardFragment (clipboardFragment.ts:839): deleteRange :857, insertTableFragment :869 (new table), insertText/insertImage | n/a | the pasted fragment needs tblPr, grid and rows; copying an incomplete table throws the coded refusal (fragmentFromPayload) |
| insertPageBreak | none | engine.insertPageBreak Actions:1345 | n/a | none |
| insertSectionBreak | none | engine.splitSection Actions:1350 | n/a | none |
| resizeObject | table-op | setColumnWidth Actions:1056; setRowHeight :1068; setColumnWidth :1084, :1085 and setTableProperties :1088 | no | the cell handle resolves through the projection (residentObjectGeometry.ts:141-163), so stale-object. If reached: missing rows leave stored undefined and setRowHeight throws (refused table-unavailable); missing grid refuses unsupported-object (:1077), not the coded reason; isUnavailableTablePayload is not used |
| moveObject | none | setImageGeometryAt Actions:1114 (image only, :685) | n/a | none |
| moveEmbed | none | engine.moveEmbed Actions:1038 (image only, :644, :652) | n/a | none |
| replaceMatches | none | proposalEngine.applyEdits replaceText Actions:1122 | n/a | a target range removes a table embed whole |
| undo | none | engine.undo then syncCellSelectionWithCaret, which turns a refusal into a cleared cell selection | n/a | a hidden caret moves next to the table |
| redo | none | engine.redo then syncCellSelectionWithCaret, which turns a refusal into a cleared cell selection | n/a | a hidden caret moves next to the table |
| addComment | none | addComment residentReviewCommands.ts:69 | n/a | none |
| updateComment | none | updateComment ReviewCommands.ts:72 | n/a | none |
| setCommentRanges | none | setCommentRanges ReviewCommands.ts:73 | n/a | none |
| removeComment | none | applyRawOps ReviewCommands.ts:74 | n/a | none |
| acceptRevision | none | acceptChange ReviewCommands.ts:75 (Rust skips unavailable tables, ops/table.rs:1695); afterwards repairCommandSelection Actions:1182 | n/a | row revisions in an unavailable table are not listed (push_table_row_changes :1609), so preflight returns unknown-revision (ReviewCommands.ts:47); a revision id spanning text and such a table resolves partly, silently (RISK 4) |
| rejectRevision | none | rejectChange ReviewCommands.ts:76 | n/a | as acceptRevision |
| propose | none | proposals registry to applyEdits (ReviewCommands.ts:58) | n/a | none |
| setProposalStates | none | proposals.setStates ReviewCommands.ts:59 (accepts and rejects through the same resolve path) | n/a | as acceptRevision |
| withdrawProposals | none | proposals.withdraw ReviewCommands.ts:60 | n/a | as acceptRevision |
| navigateCell | table-op | preflight engine.cellSelection residentCommandCells.ts:38 (unguarded). Apply: plain Tab or Shift+Tab calls only setCellSelection Cells.ts:74. createLastRow in the last cell (Tab, or Enter in the last cell) calls insertRow Cells.ts:65, then tablePayload :66 (cast, no check), then setCellSelection :74 | no | preflight rejects an incomplete payload (isUnavailableTablePayload) and reads the cell selection through readCellSelection; insertRow, the first write, is refused coded |
| selectTableCells | reads-payload | resolveLiveCell, then tableRangeFor to anchorsOf to validPayload (TableSelection.ts:99-107); apply writeCellSelection Actions:1212 (swallows errors) | n/a | rejected: invalid-cell-target |
| select | none | cellsForCaret (projection) Actions:746; writeCellSelection :1212 | n/a | none |
| navigate | reads-payload | resolveShiftArrowCellRange or cellSelectionHeadPosition to anchorsOf (TableSelection.ts:152, 255, 283) | n/a | rejected (validPayload); falls back to text navigation |
| selectPoint | reads-payload | shift-click: resolvePointerCellRange (residentCommandSelection.ts:475) to commonTableFor (projection) and anchorsOf | n/a | rejected (validPayload); an unavailable table has no laid-out cells to hit |
| extendSelection | reads-payload | as selectPoint | n/a | rejected |
| insertText | none | insertText or replaceRange Actions:1263-1265; writeCellSelection (guarded) | n/a | a range removes a table embed whole; ranges cannot span cells (single-story check, residentCommandAnchors.ts:49) |
| replaceSelection | none | as insertText | n/a | as insertText |
| deleteBackward | reads-payload | preflight tableBoundaryDeleteTarget (tableBoundaryDelete.ts:27-35 via Actions:788 and :485-493); apply deleteRange Actions:1276 | n/a | tableBoundaryDeleteTarget treats an incomplete payload as no table to enter, so Backspace never moves into a hidden cell |
| deleteForward | none | as deleteBackward | n/a | forward returns `at` for any table without reading the payload, so Delete before a table is a noop |
| paragraphBreak | none | deleteRange Actions:1306, splitParagraph :1311 | n/a | none |
| lineBreak | none | deleteRange Actions:1289, insertLineBreak :1293 | n/a | none |
| toggleMark | none | engine.toggleMark Actions:1370 | n/a | none |
| formatRange | none | engine.formatRange Actions:1371 | n/a | none |
| clearFormatting | none | engine.clearFormatting Actions:1372 | n/a | none |
| setParagraphAttrs | none | engine.setParagraphAttrs Actions:1373 | n/a | none |
| applyParagraphStyle | none | engine.applyParagraphStyle Actions:1374 | n/a | none |
| setReadOnly | none | none | n/a | none |
