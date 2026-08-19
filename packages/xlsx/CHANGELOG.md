# @betteroffice/xlsx

## 0.1.0

### Minor Changes

- 56fde13: A chart frame is now addressed by the drawing anchor it sits at rather than by the chart part behind it. Two anchors in one drawing may point at one `chart1.xml`, and both frames used to publish the same `ChartRegion.id`: hit-testing the second one and dragging it moved the first, the selection outline snapped to the first, and the workbook then refused to save at all because one part matched two charts. Each frame now carries its own id, so a drag moves the frame that was picked up and a save writes each anchor back where it sits.

  `ChartRegion.id` is `"<drawing part>#<anchor index>"`, opaque, and never a chart part path. `WorkbookHandle.moveChart`, `Workbook::move_chart` and `Op::SetChartAnchor` all take it.

  An anchor index names a position in a drawing, not an object, so an op stored against one outlives the structure it was recorded against: another editor that adds, reorders or removes an anchor renumbers every ordinal after it. `Op::SetChartAnchor` is therefore a compare-and-set — it carries `part` and `from`, the chart part and anchor the frame held when the op was recorded, beside `to` — and replay onto a drawing whose anchors have shifted is refused rather than landing on whichever frame now sits at that ordinal. The refusal is the typed `Error::ChartFrameShifted`, so a host replaying a stored log can drop that op and carry on instead of matching on prose. Two frames pinned to byte-identical anchors _and_ backed by one chart part are the only case left that the guard cannot tell apart, and they are interchangeable: same part, same rect.

  A worksheet that reaches one drawing through two relationships now walks it once; following each separately emitted every anchor twice, and the twins then shared an ordinal.

  The refusal that protects a shared chart part is unchanged in strength, only re-expressed per anchor: a save still demands that every frame a sheet carries match exactly one frame the source sheet was read with, that none be dropped, that none be carried twice, and that sheets sharing a part or a drawing agree on what it holds.

- fe1145a: A chart on a worksheet is now a selectable object instead of a picture the click falls through. Every frame publishes each chart's id, rect, clipped hit area and whether it can be repinned, and chrome resolves a press against that frame's own regions, so the answer cannot drift from the pixels. Clicking a chart selects it and outlines it; dragging it, or nudging it with the arrow keys, slides it through the op log as one undoable edit; and the moved anchor — cell and EMU offset alike — is written back into the drawing part on save, synthesising the `colOff`/`rowOff` a drawing omitted rather than saving a move that lost half of itself. Clicking off the chart restores the cell selection, and while a chart is selected the keyboard no longer reaches the cells hidden behind it.

  A chart the renderer could not draw is selectable too: it degrades to a placeholder but still occupies its space, so it stays an object that can be picked up and moved out of the way.

  Two limits worth stating. A chart pinned by an absolute anchor can be selected but not moved: its position lives in attributes the writer cannot rewrite, and the frame reports this as `movable: false` so the UI never offers the drag. And moving a chart is a standalone-session edit — chart state syncs as one blob per sheet, so a collaborative session refuses it exactly as it refuses freeze panes and hyperlinks.

  Minor rather than patch: `DisplayList.charts` changes its element contract. The elements gain `id`, `rect`, `clip` and `movable` beside `placeholder`, and lose `Eq` on the Rust side, so anything that _constructs_ one must be updated — `ChartA11yAttrs` survives only as an alias for readers. `Op::SetChartAnchor` is a new variant and breaks exhaustive matches on `Op`.

- 5989039: Dragging a chart works in a collaborative session. It used to fail outright with "structural operations are unavailable in collaborative mode", because repinning a chart sat in the structural list beside inserting a row — so the feature could not work at all in the mode the collaborative demo runs in.

  A chart anchor is not workbook structure. What is frozen is now what a chart _is_: the drawing anchor it sits at, the part behind it, the references it reads, and the shape of that anchor — its kind, its `editAs` mode, a one-cell extent, an absolute position. Where a grid-anchored chart sits travels through the shared document like any other edit, and a peer picks it up the same way it picks up a cell.

  Two things follow from letting an anchor merge, and both are worth knowing.

  An arriving anchor is judged only on what is true of it whatever grid it sits over: offsets in range, corners in order, extents positive. Whether it resolves to a rectangle with room to draw depends on the column widths it spans, and those are replicated too — so that question is settled where the chart is drawn, not where the update arrives. Collapsing the columns under a chart while someone else drags it leaves both replicas holding the same workbook; the chart simply has nowhere to render, which is what collapsing those columns asked for.

  One drawing anchor is one element however many sheets point at it, so a drag repins every sheet holding it, and a workbook read back from the shared document always shows them agreeing. Where two sheets have been written separately, the first in sheet order supplies the anchor.

  Chart state is one value per sheet, so two people dragging different charts on the same sheet keep only one of the two drags. The replicas agree on which; nothing is corrupted and no one is disconnected. Dragging the same chart at once already resolves to a single anchor.

### Patch Changes

- 692f2c7: Upgrade the XLSX raster backend to tiny-skia 0.12 with deterministic PNG encoding.
- ab39d50: The active tab now survives a round trip. Opening a workbook reads `activeTab` from the first `workbookView` rather than always starting on the first sheet, and saving writes the selection back — patched into the preserved `bookViews` where the source has one, emitted where it does not. An `activeTab` past the last sheet falls back to the first, and a save that leaves the active sheet alone still returns the source bytes untouched.
- 0f1ae6b: Classic charts are now part of the workbook model rather than an opaque preserved part. Their series, category and title references follow row and column edits, sheet renames and sheet removals the same way cell formulas do, the cached points beside each reference are regenerated from the edited workbook, anchors move and resize according to their `editAs` mode, and the chart part is patched in place so unmodelled chart markup survives. Structural edits on a workbook whose charts are all covered are no longer refused; a ChartEx part, an unclaimed chart part, a pivot-sourced or externally cached chart, a `sqref` extension reference, and a cache beside a reference that is not a direct one-dimensional range still refuse them. Charts are preserved from the package they were read with; creating one is not supported, and a chart-bearing model with no source package is refused rather than saved without its charts.
- d143a82: Worksheet charts now render in both the browser canvas and native PNG output. Two-cell, one-cell and absolute anchors follow worksheet geometry, including hidden rows and columns, viewport scrolling, partial visibility and frozen panes. Charts use the shared chart geometry, chart paths have matching fill and stroke behavior in both backends, and visible charts now carry screen-reader labels. Charts paint above cells, gridlines and proposal ghosts, while frozen-pane dividers and the editor's selection, presence and editing overlays remain above them. A frame that would exceed 65,536 chart operations is refused with a visible rendering error instead of returning a plausible-looking partial frame.

  A chart that cannot be drawn never takes the worksheet down with it. Every chart family the shared geometry draws is rendered: column, bar, line, pie, doughnut, area, scatter, bubble, radar, stock, of-pie and surface. A family outside that set, a 3-D plot, stacked or percent-stacked grouping, a secondary-axis combination and a logarithmic axis are each drawn as a neutral placeholder box in that chart's own anchor rectangle, rather than substituted with a plot that would misrepresent them. A chart part that cannot be read, or that is missing from the package, degrades the same way. The rest of the frame — cells, gridlines, ghosts and the other charts — renders normally, so one unsupported chart can no longer blank the sheet. A placeholder keeps a screen-reader label suffixed with "not shown", built from the chart's title, family and series counts when those could still be read, and its display-list entry carries a `placeholder` flag so surrounding UI can tell a drawn chart from an undrawn one. A chart whose anchor does not resolve to a rectangle on the grid is skipped, matching how the used-range calculation already treats it. Placeholders are charged against the same 65,536-operation frame budget.

  Known limits. The renderer does not yet reproduce authored chart-area and plot-area fills, gradients, pattern or picture fills, transparency, shape effects, shadows, glow, bevels, trendlines, error bars, drop lines, high-low lines, up/down bars, data tables or data-label layout. Axis titles, tick formatting and placement, custom number formats, line smoothing, marker symbols, pie rotation and authored doughnut hole size are not yet reproduced; fonts use the backend's available Calibri-compatible face, and unsupported text and shape effects are omitted.

- 64eecea: A workbook whose chart or drawing part cannot be read now opens with that chart missing, instead of the whole file being refused; cells, formulas and styles come through intact, and the part it declined to read is written back untouched on save. Such a workbook is cell-editable but structurally frozen: inserting or deleting rows and columns, and renaming or removing a sheet, are refused, because nothing can move what a part it cannot read names — the deal a chart part no sheet anchors already got. A chart part a drawing names but the package does not hold now freezes those edits too, where before it let them strand the chart silently.
- a7b8062: Rows and columns can be inserted and deleted again on a sheet a defined name points at with the dynamic-range idiom — a whole-column reference inside a call such as `SUM(Data!$A:$A)`, the range operator applied to `INDEX`, or the `Data!#REF!` Excel leaves behind. A spill (`A1#`), an implicit intersection (`@A1`) and a range written around whitespace move with the grid too. An edit whose effect on a name cannot be settled — a reference that could equally be a name of the same shape, or an unqualified one in a workbook-scope name — is still refused rather than guessed at.
- 623ee21: Rendering a sheet without an explicit range no longer fails on a workbook whose chart is parked far below the data. The frame still grows to take in a chart near the used range, but one that would push the image past the renderer's pixel caps is now left out of frame — the way a chart past the edge of an explicit viewport already is — instead of taking the whole render down with it. A render capped by `max_width` / `max_height` that used to pad out to the cap around such a chart now returns the tighter used-range frame, so its reported dimensions can be smaller than before.
- 36c87cf: Charts now follow the cells they reference. A chart used to be drawn from the value cache stored inside its chart part, so editing a referenced cell — or recalculating a formula that feeds one — left the chart plotting the numbers the file was saved with, and the stale plot survived a save and reopen. Every chart reference this engine can resolve is now read from the current workbook on each render, so a chart repaints with the cell edit in both the browser canvas and native PNG output.

  This is a rendering change only. The bytes a save writes back for an imported chart are unchanged: a chart part is still rewritten solely when its reference formulas move, and a chart that was merely rendered against live cells saves byte for byte as it was imported.

  A series is projected whole or not at all. A chart pairs a series' values with its categories slot by slot, so refreshing one while the other kept its authored values would draw a pairing no version of the file ever held. If any reference in a series cannot be resolved, that whole series keeps the values it was authored with. Series are independent of each other, so a chart can show one live series beside one authored series when only part of it is resolvable.

  Known limits. A reference the engine cannot resolve safely keeps the values the file was authored with, because a wrong number is worse than a stale one: references into another workbook, defined names, unions, three-dimensional and two-dimensional areas, references naming a sheet the workbook does not hold, `#REF!`, multi-level category caches, caches carrying extension markup or per-point number formats, and any part exceeding 256 references or 4,096 resolved cells. A chart whose data does not come from plain worksheet references at all — ChartEx, a pivot or external data source, or a filtered-series `sqref` — is never projected, since the references beside those do not mean what a plain one means. A populated cache whose cells the workbook reads as empty also keeps its authored values, so a chart whose source rows were deleted does not blank.

  A gap between plotted points is never projected. A cache omits the points a consumer cannot read, and the point list is read back by position rather than by index, so a cleared cell, text where a number was, or an `=NA()` in the middle of a range would slide every later value into an earlier slot and under another point's label. Any such range keeps the values it was authored with, and a chart therefore stops following the grid until the gap is filled. A gap after the last plotted point shifts nothing and still follows the cells.

  Two cases render and save differently. A text cache over cells that now hold numbers, booleans or errors is drawn with the live value, but a save still writes the text the file was authored with, because writing a number into a text cache would need the cell's number format applied and the engine will not guess it. Separately, a cache whose cells all read as empty keeps its authored values on screen, but if that reference is later moved by a row, column or sheet edit, the save regenerates it as genuinely empty — so a chart can be drawn with data and then saved without it. Reopening the saved file shows the empty chart; the data returns by pointing the reference back at populated cells.

- f5d1b03: Saving a workbook now preserves the parts the model does not represent — charts, drawings, pivot tables, comments, macros, custom XML and their relationships — instead of rebuilding the package and dropping them. Sheets you did not touch are copied through byte for byte, so an edit on one sheet no longer strips hidden rows, outline levels, rich inline strings or shared-formula attributes from the rest of the workbook. The stylesheet is left alone unless styles actually change, and adding a format now patches one pool entry instead of rewriting every pool. Chartsheets keep their type, freeze-pane and hyperlink edits reach both the worksheet and its relationship part, and an edited save drops the stale calculation chain so Excel recalculates on open. Cells keep the shared-string entry they were authored against, and keep it through row and column edits that move them. A new occurrence of the same text is written without borrowing an existing rich-text entry. A save is refused when removing only some duplicate entries would leave too few entries for the distinct formatting still used by cells. Print areas and print titles are rewritten by row and column edits, not left pointing at the old ranges. Chart formula references are rewritten by sheet rename and row or column edits, and their caches are regenerated when the modeled cache shape can be preserved.

  Known limits. A sheet you edit is still reserialized from the model, so unmodeled row, column and cell markup on that one sheet is lost. Autofilter, data-validation, conditional-formatting, table and sparkline ranges on an edited sheet remain at their source coordinates, as do internal hyperlink locations outside the formula parser. Editing an existing style pool entry regenerates that entry from the modeled subset. Sheet rename, sheet removal and row or column edits are refused while a pivot table or a chart part outside modeled coverage is preserved, because those references are not modeled; formulas naming a removed sheet keep that name instead of collapsing to `#REF!`. A row or column edit is also refused when a defined name aimed at that sheet is beyond the reference rewriter — a whole-row or whole-column reference nested inside a function, a structured table reference, or anything else the formula parser rejects. In a workbook with multiple sheets, the edit is refused for a workbook-scoped name with unqualified cell references because those references bind to whichever sheet uses the name. Collaborative sessions compare only the modeled workbook, so two peers holding the same cells but different macros or custom XML still accept each other as the same base. Replaying collaborative history drops which shared-string entry each cell was authored against; when an edited sheet is saved afterward, those cells are written as plain inline text rather than assigned another entry's formatting.

- 9534169: A chart or drawing part that two sheets both anchor is no longer written from whichever sheet came last. Saving such a workbook used to patch the shared part once, silently rewriting the chart the other sheet shows: inserting a row on one sheet moved its references and, on reopen, the other sheet's chart had moved with it. A part is now written only when every sheet anchoring it patches it into the same bytes, which covers the sheet a cache is rebuilt against, and the save is refused with an error naming the part and both sheets otherwise. Workbooks whose sheets want the same content out of a shared part — including every workbook that shares no part at all — save exactly as before, byte for byte.
- 148ba02: A workbook snapshot persisted by the previous release opens again. A replica bootstraps the workbook it opened into a document whose client ID is the head of the base fingerprint, and that fingerprint hashes the collaboration schema version — so raising the version for charts gave every workbook a different bootstrap identity. A snapshot an earlier release wrote no longer deduplicated against the one this build seeds: the two bases doubled up, one was tombstoned by client ID, and restoring reported that the shared workbook structure had changed or silently handed back the pristine file.

  A replica that has not been edited yet now takes a whole snapshot as its state and upgrades it in place, rather than merging it against a bootstrap it can never agree with. Where the two bootstraps do agree the snapshot and the merge describe the same document, so this is never the worse answer. The upgraded state, not the snapshot, is what peers are told about: the upgrade writes new structs, and an incremental update that later builds on them would sit unintegrated forever on a peer that never received them. What the frozen structure describes — sheet order and names, merges, freeze panes, hyperlinks and charts — must still match, disregarding the shared-type identities a replaced bootstrap changes by construction. A snapshot that fails that, or a whole document this build cannot read, is now an error rather than a silent no-op.

  A charted workbook pairs with a pre-chart snapshot again. Such a snapshot carries no chart state to disagree about — charts come from the file the replica opened, keyed to the sheet they were parsed from — so refusing the pairing only made charted workbooks the ones that could never be restored.

  Workbooks with a hidden row or column restore too. This release began modelling those as a zero dimension where earlier ones recorded nothing at all, and both dimension maps are fingerprinted, so such a workbook could not be recognised as the base its own snapshot had started from. The dimensions an earlier release would have stored are now read from the source sheet alongside the current ones and accepted as a legacy fingerprint, and restoring puts the hidden dimensions back rather than letting the row silently unhide.

  Each feature is now pinned to the schema that introduced it rather than to whichever schema is current, so the next version bump cannot reclassify the newest schema as predating charts and manufacture this same failure again.

## 0.0.8

### Patch Changes

- 4e04087: Formulas referencing defined names now resolve correctly, frozen panes render, and hyperlinks survive the round trip. The collaboration schema advances to version 5 and upgrades version 3 and 4 snapshots when read, so a client on this release cannot share a collaboration room with an older one: upgrade every peer together.
- 47c37b0: The default collaboration frame limit drops from 64 MiB to 16 MiB to match what the relay accepts and retains, and an oversized frame now raises a protocol error before it is sent instead of closing the socket. Hosts running their own relay can restore the previous ceiling with the `maxFrameBytes` option.
- 0d3baa1: Collaborative presence: remote collaborators' cell and range selections render as colored outlines with name flags, plus toolbar avatar chips; worksheets expose stable collaborative ids so presence survives sheet renames.

## 0.0.7

### Patch Changes

- 793b761: Render pending proposals as Word-style tracked changes: struck-through old
  values with a red run highlight, new values in green with a dashed underline
  and green run highlight, laid out side by side or new-over-old and following
  cell alignment. Proposal staging recalculates the formula graph and ghosts
  downstream dependents whose computed values change, proposal edits can carry
  a number format, and no-op proposals render unmarked.
- c6ad184: Add a Google Sheets-style toolbar to the XLSX editor backed by new engine
  APIs for range styling, number formats, selection-format aggregation, format
  painting, merge queries, and history state. Formatting is fully collaborative
  through a content-addressed style catalog (collaboration schema v3; v2 state
  does not migrate). Merging replaces intersecting ranges like Excel, parsing
  repairs overlapping merges in third-party files, and display-list font fields
  now serialize correctly so styled text renders with its real font, size, and
  weight.
- 793b761: Pending agent proposals render as in-cell tracked-change ghosts painted by the engine: the new value in green above the old value struck through in red, repainting immediately on propose, accept, and reject. Display-list text commands now serialize camelCase so cell fonts, sizes, and strike/underline offsets reach the canvas, and uninstalled workbook fonts fall back to sans-serif instead of the browser serif default.

## 0.0.6

### Patch Changes

- a34e721: Add deterministic Yrs replicas, bounded and validated sync-v1 exchange, a
  transport-agnostic npm collaboration provider, and React peer-update repainting.
  Collaborative sessions support nonstructural cell and dimension edits; inverse-op
  undo and redo remain disabled until a Yrs-aware undo manager can preserve
  concurrent edits.

## 0.0.5

### Patch Changes

- e8678aa: Route workbook sessions through the shared Rust facade and harden editing, calculation, and rendering limits.

## 0.0.4

### Patch Changes

- 6a1ab98: Publish the spreadsheet packages as ESM-only and load the WebAssembly core as a separate asset.

## 0.0.3

### Patch Changes

- 68d15b8: Fix `@betteroffice/xlsx-react` so its dependency on `@betteroffice/xlsx` resolves to the matching published version.
