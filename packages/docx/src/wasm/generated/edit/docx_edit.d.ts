/* tslint:disable */
/* eslint-disable */

/**
 * One yrs replica of the DOCX editing model, held for a JS host.
 *
 * Owns the [`EditingDoc`] plus the (single) JS update observer. The JS facade
 * multiplexes its own listener set over that one callback.
 */
export class EditSession {
    free(): void;
    [Symbol.dispose](): void;
    /**
     * Accepts tracked changes: pending insertions become plain content,
     * pending deletions are carried out; `pPrIns` marks clear so the split
     * stays, `pPrDel` marks join with the following paragraph, whose own
     * properties survive.
     *
     * `target_json` is either `{"revisionId": string}`, resolving that one
     * coalesced revision wherever it appears in any story, or
     * `{"story","startPara","startOffset","endPara","endOffset"}`, resolving
     * every tracked change overlapping that range regardless of revision id.
     * Receipt: `{"revisionIds": [string, …]}` — the ids actually resolved, in
     * resolution order and deduplicated. Resolving applies a revision rather
     * than authoring one, so it never stamps a new revision and ignores
     * suggesting mode. Errors on an empty range and on a `revisionId` that
     * matches nothing.
     */
    accept_change(target_json: string): string;
    /**
     * Adds a comment anchored to one or more ranges. `ranges_json`:
     * `[{"story","startPara","startOffset","endPara","endOffset"}, …]`;
     * `body_json` is any JSON value, stored as given. The anchors are sticky,
     * so they follow the text they cover, and the comment lives outside the
     * story rather than as an attribute on it. Receipt: `{"commentId"}` — a
     * freshly minted id. Errors when `ranges_json` is not an array of
     * well-formed ranges, or is empty.
     */
    add_comment(ranges_json: string, author: string, date: string, body_json: string): string;
    /**
     * Deletes one character at this session's collapsed selection and returns
     * the resulting binary `FrameDelta`. `direction` is `"backward"` or
     * `"forward"`; a surrogate pair is removed whole. At a paragraph boundary
     * this merges with the neighbouring paragraph instead.
     *
     * Errors on an unknown `direction`, when `expected_frame_epoch` is not a
     * non-negative safe integer, under the same selection and readiness
     * conditions as [`EditSession::apply_input`], and when there is no
     * character to delete in that direction (document start or end).
     */
    apply_delete(direction: string, expected_frame_epoch: number): Uint8Array;
    /**
     * Instrumented twin of [`EditSession::apply_delete`]: identical arguments,
     * result and error contract, but it also records stage timings for
     * [`EditSession::apply_input_profile_json`].
     */
    apply_delete_profiled(direction: string, expected_frame_epoch: number): Uint8Array;
    /**
     * Applies one ordinary insertion at this session's collapsed selection
     * and returns the resulting binary `FrameDelta`. The inserted text
     * inherits the formatting at the caret; selection, measurement inputs,
     * pagination checkpoints and display state all stay resident, so nothing
     * but the frame crosses the boundary.
     *
     * Errors when `text` is empty or contains `\r`/`\n`, when
     * `expected_frame_epoch` is not a non-negative safe integer, when there is
     * no resident selection or it no longer resolves, when the selection is
     * not collapsed, or when the resident layout state cannot absorb an edit
     * in that paragraph. The caller must then run the full layout path.
     */
    apply_input(text: string, expected_frame_epoch: number): Uint8Array;
    /**
     * Stage timings of the last profiled apply, in milliseconds:
     * `{"selectionMs","editMs","lowerMs","measureMs","paginateMs",
     * "displayInputMs","displayBuildMs","displayFinalizeMs","displayMs",
     * "encodeMs"}`. `{}` before the first profiled call.
     */
    apply_input_profile_json(): string;
    /**
     * Instrumented twin of [`EditSession::apply_input`]: identical arguments,
     * result and error contract, but it also records stage timings for
     * [`EditSession::apply_input_profile_json`]. Separate so the ordinary
     * input path pays no timer calls.
     */
    apply_input_profiled(text: string, expected_frame_epoch: number): Uint8Array;
    /**
     * Applies an update produced by this document's dedicated local worker.
     * The local origin lets the main replica's UndoManager retain ownership of
     * the edit; remote/collaboration updates must use `apply_update` instead.
     */
    apply_local_update(update: Uint8Array): void;
    /**
     * Writes `style_id` as the `pStyle` of every paragraph intersecting
     * `[start, end)`. Only that key changes: this boundary has no style
     * resolver, so it never fabricates the paragraph attributes or run marks
     * the style definition would imply. In suggesting mode the property
     * change is recorded as a `pPrChange` revision.
     */
    apply_paragraph_style(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number, style_id: string, author_name?: string | null, author_date?: string | null): void;
    /**
     * Applies a batch of raw story mutations in ONE transaction. Unlike every
     * other op here these carry no user intent: indices are story-global
     * UTF-16 units, not Locs, and nothing is inferred or stamped on the
     * caller's behalf. `ops_json` is an array of:
     *
     * - `{"op":"insert","index","text"?,"attrs"?}`
     * - `{"op":"delete","index","len"}`
     * - `{"op":"format","index","len","attrs"?}`
     * - `{"op":"insertEmbed","index","kind"?,"payload"?,"attrs"?}`
     * - `{"op":"setEmbedAttr","index","key","value"?}`
     * - `{"op":"setComment","id","ranges":[[start,end], …],"author"?,"date"?,"body"?}`
     * - `{"op":"removeComment","id"}`
     *
     * Each op's `index`, and each `setComment` range, is read against the
     * story state AFTER every preceding op in the batch. `attrs` and
     * `payload` are JSON objects written verbatim, so tracked-change stamps
     * travel inside `attrs`; comments are keyed by the given id and anchored
     * sticky. Errors on an unknown `"op"`, a missing or negative `"index"`,
     * or a missing required field, and leaves the story untouched — parsing
     * completes before the transaction opens.
     */
    apply_raw_ops(story: string, ops_json: string): void;
    /**
     * Applies seed raw operations with deterministic item ordering.
     */
    apply_seed_raw_ops(story: string, ops_json: string): void;
    /**
     * Applies a remote or incremental yrs v1 update. It commits without a
     * local origin, so this replica's undo manager never claims it. Errors on
     * a malformed update.
     */
    apply_update(update: Uint8Array): void;
    /**
     * [`EditSession::apply_update`] plus an inference about where the peer
     * that authored it is typing, for caret presence. Returns
     * `{"clientId","story","paraId","endOffset"}` for the end of that peer's
     * last text insertion, or `"null"` when the update carries work from more
     * than one client, ends in a non-text insertion, or inserts nothing.
     * Errors on a malformed update.
     */
    apply_update_with_inference(update: Uint8Array): string;
    /**
     * Display-only input JSON in, one binary `FrameDelta` v1 out (exposed as
     * a transferable `Uint8Array`). `expected_frame_epoch` is the epoch of the
     * frame the caller currently holds; pass `0` for the first frame. A
     * mismatch makes the engine emit a full frame instead of a delta. Errors
     * unless the epoch is a non-negative safe integer, and on build failure.
     */
    build_display_list_frame(input: string, expected_frame_epoch: number): Uint8Array;
    /**
     * `{ measured, options, layout }` JSON in, `DisplayList` JSON out, built
     * against the same resident font store this session measures with.
     */
    build_display_list_json(input: string): string;
    /**
     * Whether [`EditSession::redo`] would reapply something.
     */
    can_redo(): boolean;
    /**
     * Whether [`EditSession::undo`] would revert something.
     */
    can_undo(): boolean;
    /**
     * The current local cell selection as a [`TableRange`], or `"null"`
     * before one is set. An endpoint whose cell was deleted clamps to a
     * surviving nearby cell. Errors when the table itself no longer resolves.
     */
    cell_selection(): string;
    /**
     * Removes the authored `value` from the content-control embed carrying
     * `embed_id`, leaving the control itself in place. Errors when no embed
     * has that id.
     */
    clear_content_control_value(embed_id: string): void;
    /**
     * Clears every direct run-formatting attribute over `[start, end)`.
     * Protected attributes — hyperlinks and tracked-change stamps — survive,
     * as do paragraph properties.
     */
    clear_formatting(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number): void;
    /**
     * Drops every registered measurement font (ids restart at zero) and
     * invalidates the retained paragraph measurement templates, so the next
     * edit must pass back through the full layout path.
     */
    clear_measure_fonts(): void;
    /**
     * Stops queueing and discards anything not yet drained.
     */
    clear_update_event_observation(): void;
    /**
     * Drops the update observer registered by [`EditSession::set_update_observer`].
     */
    clear_update_observer(): void;
    /**
     * This replica's client id, as passed to the constructor.
     */
    client_id(): number;
    /**
     * Adds a story holding one paragraph with `initial_text` (which must not
     * contain paragraph breaks), `p_style` and `alignment`. Receipt:
     * `{"paraId"}` — the paragraph ending at the story's pilcrow. Errors when
     * the story id already exists.
     */
    create_story(story_id: string, initial_text: string, p_style: string, alignment: string): string;
    /**
     * Deletes every column the [`TableRange`] `range_json` covers, along with
     * the stories of the cells removed. Always a plain local edit.
     */
    delete_column(range_json: string): string;
    /**
     * Deletes `[start, end)`. Because a range crossing a paragraph boundary
     * includes the boundary pilcrow, a plain delete also merges those
     * paragraphs. Suggesting mode removes nothing and stamps the content
     * `del` instead. Receipt: `{"revisionId": string|null}`.
     */
    delete_range(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number, author_name?: string | null, author_date?: string | null): string;
    /**
     * Deletes every row the [`TableRange`] `range_json` covers. In suggesting
     * mode the rows are marked `trDel` instead of removed.
     */
    delete_row(range_json: string, author_name?: string | null, author_date?: string | null): string;
    /**
     * Removes one complete story and its content. Errors on an unknown story.
     */
    delete_story(story_id: string): void;
    /**
     * Removes the table `table_json` ([`TableLocator`]) names plus every one
     * of its reachable cell stories. Always a plain local edit; the receipt
     * has `deletedTable: true`.
     */
    delete_table(table_json: string): string;
    /**
     * Region-aware hit test against the resident display list, so no
     * display-list JSON crosses the boundary. `x`/`y` are page-local px.
     * Returns
     * `{"region":"body"|"header"|"footer"|"footnote"|"endnote","rId"?,`
     * `"noteId"?,"pos":n|null,"target":"text"|"image"|"none"}`, or `"null"`
     * for a page index outside the frame. A header/footer `pos` refers to
     * that part's own document and a note `pos` to the `fn:{id}` / `en:{id}`
     * story, never the body; `target` is what sits under the point, which is
     * what a pointer cursor needs.
     */
    display_hit_test_regions_json(page_index: number, x: number, y: number): string;
    /**
     * Highlight rectangles for a body document range against the resident
     * display list: a JSON array of `{pageIndex,x,y,width,height}` in
     * page-local px, one entry per page the range touches. Body positions
     * only. Errors when no display list is resident.
     */
    display_range_rects_json(from: number, to: number): string;
    /**
     * Same rectangles as [`EditSession::display_range_rects_json`], scoped to
     * a region. `region` is `"body"`, `"header"`, `"footer"`, `"footnote"` or
     * `"endnote"`; `part_id` names one header/footer part (an empty string
     * matches any) or the note whose story the positions belong to.
     * `from`/`to` are positions in THAT region's document, and a
     * header/footer part paints on every page carrying it, so the result
     * holds one rect set per such page, each tagged with its `pageIndex`.
     * Errors on an unknown `region` and when no display list is resident.
     */
    display_range_rects_region_json(region: string, part_id: string, from: number, to: number): string;
    /**
     * Nearest caret position on the adjacent visual line. `direction` is
     * `"up"` or `"down"`; `goal_x` is the page-local x the caret is trying to
     * hold across successive moves (a non-finite value falls back to the
     * caret's own x). Returns `{"position","goalX"}`, or `"null"` when there
     * is no line in that direction. Errors on an unknown `direction` and when
     * no display list is resident.
     */
    display_vertical_move_json(position: number, direction: string, goal_x: number): string;
    /**
     * Pops the oldest queued transaction, in arrival order. Byte 0 is `1`
     * when the transaction had no local origin and `0` otherwise; the rest is
     * the v1 update. Empty when the queue is drained or observation never
     * started.
     */
    drain_update_event(): Uint8Array;
    /**
     * The yrs v1 update carrying everything this replica has that the peer
     * described by `remote_state_vector` does not. Errors on a malformed
     * state vector.
     */
    encode_diff(remote_state_vector: Uint8Array): Uint8Array;
    /**
     * The full document state as one yrs v1 update. Hand it to
     * [`EditSession::load`] on a fresh replica to reproduce this document.
     */
    encode_state(): Uint8Array;
    /**
     * This replica's yrs v1 state vector — what a peer sends so
     * [`EditSession::encode_diff`] can compute the update it is missing.
     */
    encode_state_vector(): Uint8Array;
    /**
     * Encodes one Loc as opaque sticky-position bytes that keep pointing at
     * the same content as the story is edited. Publish them over an awareness
     * transport and resolve them with
     * [`EditSession::resolve_sticky_position`].
     */
    encode_sticky_position(story: string, para_id: string, offset: number): Uint8Array;
    /**
     * This peer's selection in transportable form:
     * `{"story","anchor":[byte, …],"head":[byte, …]}` with sticky-encoded
     * endpoints, or `"null"` before a selection is set. A peer resolves it
     * with [`EditSession::resolve_encoded_selection`].
     */
    encoded_selection(): string;
    /**
     * Applies a set-valued, tri-state inline formatting delta over
     * `[start, end)` in one transaction. An omitted key keeps the current
     * value, `null` clears it, and any other value sets it. `delta_json`:
     *
     * ```json
     * {
     *   "bold": true, "italic": true,
     *   "underline": true | {"style"?: string, "color"?: string},
     *   "strike": true | {"double"?: boolean},
     *   "color": {"rgb": string} | {"themeColor": string},
     *   "highlight": string,
     *   "fontSize": 12,
     *   "fontFamily": {"ascii": string, "hAnsi"?: string},
     *   "other": {"anyAttr": value}
     * }
     * ```
     *
     * `false` clears `bold`, `italic`, `underline` and `strike`. `color`
     * takes exactly one of `rgb` or `themeColor`. `other` writes arbitrary
     * run attributes, with `null` removing one. Always a plain local edit, so
     * there is no receipt. Errors when a key carries a type not listed here.
     */
    format_range(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number, delta_json: string): void;
    /**
     * Inserts a column left (`after = false`) or right (`after = true`) of
     * the cell `at_json` ([`CellLoc`]) names. Always a plain local edit.
     */
    insert_column(at_json: string, after: boolean): string;
    /**
     * Inserts one inline image embed at `(story, para_id, offset)`.
     * `payload_json` is the image's authored payload object, stored as given.
     * The embed occupies one story unit. Receipt:
     * `{"revisionId": string|null}`. Errors when the payload is not an
     * object.
     */
    insert_image(story: string, para_id: string, offset: number, payload_json: string, author_name?: string | null, author_date?: string | null): string;
    /**
     * Inserts a page-break embed at `(story, para_id, offset)`, occupying one
     * story unit. Always a plain local edit.
     */
    insert_page_break(story: string, para_id: string, offset: number): void;
    /**
     * Inserts a row above (`after = false`) or below (`after = true`) the
     * cell `at_json` names. `at_json` is a [`CellLoc`].
     */
    insert_row(at_json: string, after: boolean, author_name?: string | null, author_date?: string | null): string;
    /**
     * Inserts a section-break embed at `(story, para_id, offset)`.
     * `break_type` must be `"nextPage"`, `"continuous"`, `"oddPage"` or
     * `"evenPage"`; anything else errors. Always a plain local edit.
     */
    insert_section_break(story: string, para_id: string, offset: number, break_type: string): void;
    /**
     * Inserts a `rows` x `columns` table at `(story, para_id, offset)`,
     * creating one story per cell. Errors when either dimension is zero.
     */
    insert_table(story: string, para_id: string, offset: number, rows: number, columns: number, author_name?: string | null, author_date?: string | null): string;
    /**
     * Inserts `text` at `(story, para_id, offset)`. It must contain no
     * paragraph or line breaks, and it inherits the formatting at the
     * insertion point. Receipt: `{"revisionId": string|null}` — non-null in
     * suggesting mode, where the text is stamped `ins` and coalesces into an
     * adjacent insertion by the same author rather than opening a second
     * revision.
     */
    insert_text(story: string, para_id: string, offset: number, text: string, author_name?: string | null, author_date?: string | null): string;
    /**
     * Inserts a watermark embed at `(story, para_id, offset)`, its payload
     * taken verbatim from the `watermark_json` object. Always a plain local
     * edit. Errors when the JSON is not an object.
     */
    insert_watermark(story: string, para_id: string, offset: number, watermark_json: string): void;
    /**
     * `{ measured, options }` JSON in, `Layout` JSON out. Both the measured
     * input and the resulting layout are retained for the resident edit path.
     * Errors on unparseable input or on layout failure.
     */
    layout_document_json(input: string): string;
    /**
     * Region-layout input JSON in, a paginated envelope with section and page
     * regions already composed as JSON out — ready for the display builder
     * with no host-side layout mutation. Retains the pass for resident edits.
     */
    layout_document_with_regions_json(input: string): string;
    /**
     * Region-layout input JSON in, the font families and sizes that input
     * needs as JSON out, so the host can register fonts before laying out.
     */
    layout_font_requirements_json(input: string): string;
    /**
     * Every pending tracked change across all stories, in deterministic
     * story-then-position order:
     * `[{"revisionId","author","date","kind","story","preview",
     * "range":{"story","start":{"paraId","offset"},"end":{…}}}, …]`. `kind`
     * is one of `"insertion"`, `"deletion"`, `"pPrIns"`, `"pPrDel"`,
     * `"pPrChange"`, `"trIns"`, `"trDel"`, `"tableIns"` or `"tableDel"`.
     * `preview` holds the first few characters of the affected text, and is
     * empty for every structural kind, which covers no text.
     */
    list_revisions(): string;
    /**
     * Hydrates this replica from an encoded yrs v1 update, typically another
     * replica's [`EditSession::encode_state`] output. Identical to
     * [`EditSession::apply_update`]; the separate name marks the initial-load
     * call site. Errors on a malformed update.
     */
    load(update: Uint8Array): void;
    /**
     * Seeds stories from JSON:
     * `[{"storyId","paragraphs":[{"text","pStyle"?,"alignment"?}, …]}, …]`.
     * `pStyle` defaults to `"Normal"` and `alignment` to `"left"`; paragraph
     * text must not contain paragraph breaks. Receipt:
     * `{storyId: [paraId, …]}` with each story's paragraphs in document
     * order. Errors when a story entry has no `storyId`, no `paragraphs`
     * array, or an empty one, and when a story id already exists.
     */
    load_json(stories_json: string): string;
    /**
     * `{"start","end"}` — the paragraph's span in story-global UTF-16 units.
     * `end` is the index of its own pilcrow, so `end - start` is the
     * paragraph length and the upper bound of a Loc `offset` in it. Errors
     * when the paragraph is not in `story`.
     */
    locate_paragraph(story: string, para_id: string): string;
    /**
     * Re-parses the DOCX bytes retained by the last
     * [`EditSession::open_docx`] and returns the COMPLETE package envelope as
     * JSON, or `None` when no DOCX has been opened. Unlike `open_docx` this
     * keeps every part; it reflects the source file, not later edits.
     */
    materialize_docx(): string | undefined;
    /**
     * Measurement input JSON in, `ParagraphExtent` JSON out. Also records the
     * paragraph's immutable width/font envelope under its stable block id, so
     * a later resident edit re-measures only the changed block. Errors with
     * the engine's message for input it cannot measure.
     */
    measure_paragraph_json(input: string): string;
    /**
     * Merges the rectangle the [`TableRange`] `range_json` covers into its
     * top-left cell, whose story survives; the other cells' stories are
     * deleted. Always a plain local edit. Errors when the range covers fewer
     * than two cells or has no top-left anchor cell.
     */
    merge_cells(range_json: string): string;
    /**
     * Merges `para_id` with the FOLLOWING paragraph by deleting (plain) or
     * `del`- and `pPrDel`-marking (suggesting) its pilcrow. On a plain merge
     * the survivor adopts the deleted mark's properties and paraId, so the
     * earlier paragraph's identity wins. Receipt:
     * `{"revisionId": string|null}`. Errors on the story's final paragraph,
     * which has no following paragraph to merge with.
     */
    merge_paragraphs(story: string, para_id: string, author_name?: string | null, author_date?: string | null): string;
    /**
     * Creates a replica. The host allocates `client_id` (a random 32-bit id
     * is fine) and must keep it unique across the replicas that will merge.
     * Errors unless it is a non-negative safe integer.
     */
    constructor(client_id: number);
    /**
     * Parses a DOCX package, optionally seeds its editable stories into this
     * replica, and retains the source bytes for
     * [`EditSession::materialize_docx`].
     *
     * Returns `{"envelope","referencedFonts":[string, …]}`. The envelope is
     * the parsed package with the parts the host does not need stripped —
     * body content, header/footer and note content, numbering, media and
     * charts are emptied, section entries keep only their properties — so
     * styles, theme, settings, fonts and relationships still cross while the
     * bulk of the document stays in Rust. Errors on bytes that are not a
     * readable DOCX.
     */
    open_docx(bytes: Uint8Array, seed_stories: boolean): string;
    /**
     * One glyph outline from this session's resident font store:
     * `{"upem":n,"cmds":[{"t":"M"|"L"|"Q"|"C"|"Z", …}]}` — commands in font
     * units, y-up. `font_id` comes from
     * [`EditSession::register_measure_font`] and `glyph_id` from shaping.
     * Errors when the glyph cannot be extracted.
     */
    outline_glyph_json(font_id: number, glyph_id: number): string;
    /**
     * Compact paragraph-position projection built in one story traversal:
     * `[{"paraId","length"}, …]` in document order, where `length` counts the
     * UTF-16 text and inline embed units before that paragraph's pilcrow —
     * exactly the `offset` domain of a Loc in it. One call replaces crossing
     * the boundary once per paragraph. Errors on an unknown story.
     */
    paragraph_spans(story: string): string;
    /**
     * `[{"paraId","text","properties"}, …]` in document order. `text` is the
     * paragraph's plain text without its pilcrow; `properties` is the
     * pilcrow's authored property map (`pStyle`, `alignment` and whatever
     * else has been set on it). Errors on an unknown story.
     */
    paragraphs(story: string): string;
    /**
     * Reapplies the latest locally undone transaction and reports whether
     * anything was reapplied.
     */
    redo(): boolean;
    /**
     * Current local redo stack size. Zero before a story starts tracking.
     */
    redo_depth(): number;
    /**
     * Registers raw sfnt bytes in this session's resident measurement store
     * and returns the font id that measurement and display inputs reference.
     * Errors on bytes the font parser rejects.
     */
    register_measure_font(bytes: Uint8Array): number;
    /**
     * Rejects tracked changes — the inverse of
     * [`EditSession::accept_change`]: pending insertions roll back, pending
     * deletions restore their text; `pPrIns` marks join back with the
     * following paragraph, `pPrDel` marks clear so the split stays, and a
     * `pPrChange` restores the paragraph's previous properties. Same target,
     * receipt and error contract.
     */
    reject_change(target_json: string): string;
    /**
     * Replaces `[start, end)` with `text` in one transaction. The inserted
     * text adopts the first replaced unit's formatting; in suggesting mode
     * the deletion and the insertion share one revision id. Receipt:
     * `{"revisionId": string|null}`.
     */
    replace_range(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number, text: string, author_name?: string | null, author_date?: string | null): string;
    /**
     * `{"frameEpoch", "caretRect": {…}|null}` for the session's own collapsed
     * body selection. `caretRect` is null whenever there is no selection, the
     * selection is not a collapsed body caret, or the retained layout has no
     * geometry for it. `frameEpoch` identifies the frame the rect belongs to.
     */
    resident_caret_snapshot_json(): string;
    /**
     * Where a comment's sticky anchors currently sit:
     * `[{"story","start","end"}, …]`, one entry per anchored range, in
     * story-global UTF-16 units. Errors on an unknown comment id and when an
     * anchor no longer resolves.
     */
    resolve_comment(comment_id: string): string;
    /**
     * Resolves another peer's encoded selection against this replica's
     * current state: `{"anchor":{"story","paraId","offset"},"head":{…}}`.
     * Errors on malformed bytes or an endpoint that no longer resolves.
     */
    resolve_encoded_selection(story: string, anchor: Uint8Array, head: Uint8Array): string;
    /**
     * Resolves sticky bytes from [`EditSession::encode_sticky_position`] back
     * to `{"story","paraId","offset"}`. Errors when the bytes are malformed
     * or the position no longer resolves in `story`.
     */
    resolve_sticky_position(story: string, position: Uint8Array): string;
    /**
     * [`EditSession::open_docx`] with seeding always on.
     */
    seed_from_docx(bytes: Uint8Array): string;
    /**
     * This peer's current selection as `{"anchor":{"story","paraId","offset"},
     * "head":{…}}`, or `"null"` before [`EditSession::set_selection`] is
     * called. Errors when an endpoint no longer resolves.
     */
    selection(): string;
    /**
     * Aggregated toolbar and accessibility state over `[start, end)`:
     *
     * ```json
     * {
     *   "bold": true | false | "mixed", "italic": …, "underline": …, "strike": …,
     *   "fontFamily": string|null, "fontSize": number|null, "color": string|null,
     *   "paraId": string, "styleId": string|null, "alignment": string|null,
     *   "paragraphProperties": {…},
     *   "hasSelection": bool, "isMultiParagraph": bool, "inTable": bool,
     *   "isSingleEmbed": bool, "embedKind": string|null, "isImage": bool,
     *   "inInsertion": bool, "inDeletion": bool
     * }
     * ```
     *
     * A toggle mark is `"mixed"` when the range disagrees about it; a value
     * mark is `null` when it is absent OR not uniform, so the two cases are
     * not distinguished. `isImage` is `embedKind == "image"`, and
     * `inInsertion`/`inDeletion` report whether the range sits inside a
     * pending tracked change.
     */
    selection_context(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number): string;
    /**
     * Merges the sides of the JSON object `borders_json` into every selected
     * cell's `tcPr.borders`; `insideH`/`insideV` resolve to the physical edges
     * interior to the selection. `range_json` is a [`TableRange`]. Always a
     * plain local edit.
     */
    set_cell_borders(range_json: string, borders_json: string): string;
    /**
     * Stores a rectangular anchor-cell to head-cell selection outside the yrs
     * document. `range_json` is a [`TableRange`]:
     * `{"anchor":{"story","tableIndex","row","column"},"head":{…}}`. The table
     * embed is held by a sticky index and each endpoint by its cell story's
     * stable identity, so the selection survives unrelated edits and follows
     * inserted or deleted rows and columns. Errors when the two endpoints are
     * not in the same table, or when either cell does not resolve.
     */
    set_cell_selection(range_json: string): void;
    /**
     * Sets every selected cell's background to `color` (hex, with or without
     * a leading `#`), or clears it when `color` is absent. `range_json` is a
     * [`TableRange`]. Always a plain local edit.
     */
    set_cell_shading(range_json: string, color?: string | null): string;
    /**
     * Merges the JSON object `patch_json` into every selected cell's `tcPr`;
     * a `null` value removes that key. `range_json` is a [`TableRange`].
     * Always a plain local edit. Errors when the patch touches `rowspan`,
     * `colspan`, `gridSpan` or `vMerge`, which only merge and split may set.
     */
    set_cell_text_format(range_json: string, patch_json: string): string;
    /**
     * Sets the grid width, in twips, of the column holding the cell `at_json`
     * ([`CellLoc`]) names. Always a plain local edit. Errors unless
     * `width_twips` is finite and positive.
     */
    set_column_width(at_json: string, width_twips: number): string;
    /**
     * Sets the authored `value` (any JSON) on the content-control embed
     * carrying `embed_id`, searching every story. Errors when no embed has
     * that id.
     */
    set_content_control_value(embed_id: string, value_json: string): void;
    /**
     * Sets the authored `value` on the content-control embed at
     * `(story, para_id, offset)` — the way to reach a control with no
     * authored `w:id` or tag, which
     * [`EditSession::set_content_control_value`] cannot address. Errors when
     * that position holds no embed.
     */
    set_content_control_value_at(story: string, para_id: string, offset: number, value_json: string): void;
    /**
     * Sets or clears the hyperlink attribute over `[start, end)`.
     * `hyperlink_json` is `{"href", "tooltip"?, "rId"?}` or `null` to unlink.
     * The attribute is protected: ordinary formatting ops cannot write or
     * erase it, only this one. Errors when the JSON is neither an object nor
     * `null`.
     */
    set_hyperlink(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number, hyperlink_json: string): void;
    /**
     * Writes geometry fields onto the image embed carrying `embed_id` in one
     * transaction. Every key of the `geometry_json` object becomes a payload
     * entry, with `null` clearing it; the entries of a nested `"other"`
     * object are flattened into the payload alongside them. Errors when
     * `geometry_json` or its `"other"` is not an object, and when no embed
     * has that id.
     */
    set_image_geometry(embed_id: string, geometry_json: string): void;
    /**
     * Sets one paragraph property to any JSON value on `para_id`'s pilcrow,
     * searching every story. Unlike
     * [`EditSession::set_paragraph_attrs`] this writes a single key and never
     * records a revision. Errors when the paragraph is unknown and when `key`
     * names schema-managed identity (`paraId` or the embed discriminator).
     */
    set_paragraph_attr(para_id: string, key: string, value_json: string): void;
    /**
     * Applies a tri-state paragraph-property delta to every paragraph
     * intersecting `[start, end)` in one transaction. An omitted key keeps
     * the current value and `null` clears it. `attrs_json` recognises
     * `alignment`, `lineSpacing`, `lineSpacingRule`, `spaceBefore`,
     * `spaceAfter`, `indentLeft`, `indentRight`, `indentFirstLine`,
     * `hangingIndent`, `bidi`, `tabs`
     * (`[{"position":number,"alignment":string,"leader"?:string}, …]`) and
     * `defaultTextFormatting` (an object of run defaults). Any other key —
     * whether written at the top level or nested under `"other"` — is stored
     * as an opaque paragraph property. Spacing and indents are authored OOXML
     * units (twips, line-spacing units), never pixels. In suggesting mode a
     * change is recorded as a `pPrChange` revision. Errors when a recognised
     * key carries the wrong type, and when a key names schema-managed
     * identity such as `paraId`.
     */
    set_paragraph_attrs(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number, attrs_json: string, author_name?: string | null, author_date?: string | null): void;
    /**
     * Stores this peer's anchor and head as sticky positions, replacing any
     * previous selection. Both endpoints must lie in `story`. The positions
     * live outside the yrs document, so they are never serialized as content
     * or carried in an update. `Assoc::After` makes a collapsed caret advance
     * with text inserted at it.
     */
    set_selection(story: string, anchor_para: string, anchor_offset: number, head_para: string, head_offset: number): void;
    /**
     * Sets the preferred width, in twips, of the table `table_json`
     * ([`TableLocator`]) names. Always a plain local edit. Errors unless
     * `width_twips` is finite and positive.
     */
    set_table_width(table_json: string, width_twips: number): string;
    /**
     * Subscribes `callback(update: Uint8Array, isRemote: 0|1)` to every
     * committed transaction. `update` is v1-encoded — feed it straight to
     * [`EditSession::apply_update`] on a peer — and is copied out of wasm
     * memory, so JS may hold it across later edits. The second argument is
     * `1` when the transaction had no local origin. One observer per session;
     * a second call replaces the first, and a throwing callback is ignored.
     */
    set_update_observer(callback: Function): void;
    /**
     * Splits the cell `at_json` ([`CellLoc`]) covers into a `rows` x
     * `columns` grid; the original story stays in the top-left slot and every
     * other slot gets a fresh one-paragraph cell story. Omitting both
     * dimensions unmerges the cell into the slots it already covers, which
     * then requires a merged cell. Always a plain local edit. Errors on a
     * zero dimension or a grid smaller than the cell already covers.
     */
    split_cell(at_json: string, rows?: number | null, columns?: number | null): string;
    /**
     * Splits a paragraph at `(story, para_id, offset)` by inserting one
     * pilcrow. The FIRST half keeps the original paraId and the second is
     * re-minted. A split at the paragraph end leaves the empty second half
     * with only the inherited property subset; a mid-paragraph split keeps
     * its properties. Paragraph borders are cleared either way. Suggesting
     * mode stamps the new pilcrow `ins` and `pPrIns`. Receipt:
     * `{"firstParaId","secondParaId","revisionId": string|null}`.
     */
    split_paragraph(story: string, para_id: string, offset: number, author_name?: string | null, author_date?: string | null): string;
    /**
     * Starts queueing committed transactions for
     * [`EditSession::drain_update_event`] instead of pushing them through a
     * callback. Idempotent while observation is running.
     */
    start_update_event_observation(): void;
    /**
     * The story's `canonical-stream-v1` FNV-1a checksum (see
     * [`crate::canonical`]) as a DECIMAL STRING, because a u64 exceeds the
     * JavaScript safe-integer range. Two stories with the same authored
     * content share a checksum even when their paragraph and comment ids
     * differ. Errors on an unknown story.
     */
    story_checksum(story: string): string;
    /**
     * Every story id in the document, sorted so the order is stable across
     * replicas.
     */
    story_ids(): string[];
    /**
     * Story length in UTF-16 units, every embed (pilcrows included) counting
     * as one. Errors on an unknown story.
     */
    story_len(story: string): number;
    /**
     * The story as an ordered run of formatted segments — the same view
     * lowering reads:
     *
     * - `{"kind":"text","text","attributes"}`
     * - `{"kind":"pilcrow","paraId","properties","attributes"}`
     * - `{"kind":"embed","embedKind","payload","attributes"}`
     *
     * A text segment covers one maximal run of identically formatted
     * characters; each pilcrow and embed is its own segment worth one unit.
     * `attributes` holds the segment's run marks together with any `ins`/`del`
     * tracked-change stamps. Errors on an unknown story.
     */
    story_segments(story: string): string;
    /**
     * Applies one run mark over `[start, end)`. `mark_json`:
     * `{"type":"bold"|"italic"|"underline"|"strike"|"superscript"|"subscript"} |
     * {"type":"fontFamily"|"color","value":string} |
     * {"type":"fontSize","value":number}`. The six boolean types TOGGLE —
     * they turn on unless the whole range already carries the mark — while
     * font family, size and color SET. Formatting is always a plain local
     * edit, never a tracked change, so there is no receipt. Errors on an
     * unknown `"type"` or a missing/mistyped `"value"`.
     */
    toggle_mark(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number, mark_json: string): void;
    /**
     * Starts local undo tracking for a structural table edit in `story`.
     * Besides the parent story, which owns the table embed, this widens the
     * scope to the stories root so undo and redo also remove and restore the
     * cell stories the edit created or destroyed. Tracked separately from
     * [`EditSession::track_undo`] on the same story, so switching between
     * them starts a fresh history. Errors on an unknown story.
     */
    track_table_undo(story: string): void;
    /**
     * Starts local-origin undo tracking for one story, replacing any scope
     * already tracked (and its history). Call this after import or seeding but
     * before the first edit, so the initial document is not an undo step.
     * Re-tracking the same story is a no-op that preserves the history.
     * Errors on an unknown story.
     */
    track_undo(story: string): void;
    /**
     * Reverts the latest local-origin transaction and reports whether
     * anything was reverted. Remote and system transactions are excluded by
     * the manager's tracked-origin policy; `false` before a story is tracked.
     */
    undo(): boolean;
    /**
     * Current local undo stack size. Zero before a story starts tracking.
     */
    undo_depth(): number;
    /**
     * Lowers one story to a `LayoutBlock[]` JSON array — the block, run and
     * table vocabulary the layout engine consumes. `env_json` supplies the
     * document-level values lowering cannot read off the story:
     * `{"themeColors":{slot: hex},"defaultTabStopTwips":number|null,
     * "pageContentHeight":number|null,"numericIds":{yrsId: number}}`, all
     * optional. Errors when the story does not end in a pilcrow, holds a
     * malformed table, references itself through a cell story, or contains an
     * embed lowering does not support.
     */
    yrs_blocks_for_story(story: string, env_json: string): string;
}

/**
 * wasm compatibility wrapper. Resident engine users call
 * [`build_display_list_value`] and keep the typed result.
 */
export function build_display_list_json(input: string): string;

/**
 * Drop every registered measurement font (ids restart at 0). Callers must
 * re-register before the next `measure_paragraph_json`.
 */
export function clear_measure_fonts(): void;

/**
 * wasm wrapper over [`session::close_display_list`]: drop a handle so its
 * parsed display list is freed. Idempotent.
 */
export function close_display_list(handle: number): void;

/**
 * wasm wrapper over [`hit::hit_test_json`]: display-list JSON + page-local
 * point in, document position (or `null`) as JSON out.
 */
export function hit_test_json(display_list: string, page_index: number, x: number, y: number): string;

/**
 * wasm wrapper over [`session::hit_test_regions_by_handle`]: region-aware hit
 * test against a stored display list. `Err` on an unknown/closed handle so the
 * caller can fall back to [`hit_test_regions_json`].
 */
export function hit_test_regions_by_handle(handle: number, page_index: number, x: number, y: number): string;

/**
 * wasm wrapper over [`hit::hit_test_regions_json`]: region-aware hit test —
 * `{"region":"body"|"header"|"footer"|"footnote"|"endnote","rId"?,"noteId"?,`
 * `"pos":n|null,"target":"text"|"image"|"none"}` (or `"null"` for an
 * out-of-range page). The plain `hit_test_json` export stays body-only.
 */
export function hit_test_regions_json(display_list: string, page_index: number, x: number, y: number): string;

/**
 * wasm panics abort, so a trap reaches JS as a bare `RuntimeError:
 * unreachable executed`. Log the panic first so it stays diagnosable. Runs on
 * module init for every wasm core that links this crate (layout and edit).
 */
export function install_panic_hook(): void;

/**
 * wasm wrapper over [`layout_to_json`].
 */
export function layout_document_json(input: string): string;

/**
 * Measures a paragraph: measurement input JSON in, `ParagraphExtent` JSON
 * out. An `Err` whose message starts with `"UNSUPPORTED"` means the caller
 * must fall back to browser measurement for that block.
 */
export function measure_paragraph_json(input: string): string;

/**
 * wasm wrapper over [`session::open_display_list`]: parse a display list once
 * and return a handle the by-handle query exports reuse (no per-query
 * re-parse). The caller frees it with [`close_display_list`]. `Err` on
 * malformed JSON — the caller then stays on the JSON-arg path.
 */
export function open_display_list(display_list: string): number;

/**
 * wasm wrapper over [`ooxml_text::FontStore::outline_glyph_json`]: the outline
 * of a registered font's glyph, in font design units, as JSON:
 * `{"upem":2048,"cmds":[{"t":"M","x":..,"y":..},{"t":"L","x":..,"y":..},
 * {"t":"Q","cx":..,"cy":..,"x":..,"y":..},
 * {"t":"C","c1x":..,"c1y":..,"c2x":..,"c2y":..,"x":..,"y":..},{"t":"Z"}]}`.
 * The canvas caches this per `(fontId, glyphId)` and scales by `size/upem`,
 * flipping y at draw time. `cmds` is empty for a blank glyph (space).
 */
export function outline_glyph_json(font_id: number, glyph_id: number): string;

/**
 * Wasm control-plane entry: safe ZIP -> bounded XML -> typed relationships.
 */
export function parse_docx_relationships(data: Uint8Array): string;

/**
 * Parses an S2 package projection.
 */
export function parse_docx_s2(data: Uint8Array): string;

/**
 * Parses an S3 package projection.
 */
export function parse_docx_s3(data: Uint8Array): string;

/**
 * Parses an S4 package projection.
 */
export function parse_docx_s4(data: Uint8Array): string;

/**
 * Parses an S5 package projection.
 */
export function parse_docx_s5(data: Uint8Array): string;

/**
 * Parses an S6 package projection.
 */
export function parse_docx_s6(data: Uint8Array): string;

/**
 * Parses an S7 package projection.
 */
export function parse_docx_s7(data: Uint8Array): string;

/**
 * Parses an S8 package projection.
 */
export function parse_docx_s8(data: Uint8Array): string;

/**
 * Parses the full document wire in one package pass.
 */
export function parse_docx_s9(data: Uint8Array, options_json: string): string;

/**
 * Focused wasm leaf used by hostile-input and facade tests.
 */
export function parse_relationships_xml(xml: Uint8Array, part_path: string): string;

/**
 * wasm wrapper over [`session::range_rects_by_handle`]: range rects against a
 * stored display list. `Err` on an unknown/closed handle so the caller can
 * fall back to [`range_rects_json`].
 */
export function range_rects_by_handle(handle: number, from: number, to: number): string;

/**
 * wasm wrapper over [`hit::range_rects_json`]: display-list JSON + document range
 * in, JSON array of page-local rects out.
 */
export function range_rects_json(display_list: string, from: number, to: number): string;

/**
 * wasm wrapper over [`session::range_rects_region_by_handle`]: region-aware
 * range rects against a stored display list. `region` is
 * `"body" | "header" | "footer" | "footnote" | "endnote"`; `part_id` scopes
 * header/footer to one HF part and names the note id for a note region.
 * `Err` on an unknown/closed handle so the caller can fall back to
 * [`range_rects_region_json`].
 */
export function range_rects_region_by_handle(handle: number, region: string, part_id: string, from: number, to: number): string;

/**
 * wasm wrapper over [`hit::range_rects_region_json`]: region-aware range rects.
 * `region` is `"body" | "header" | "footer" | "footnote" | "endnote"`;
 * `part_id` scopes a header/footer to one HF part (empty for body / match-any)
 * and names the note id for a note region. The `from`/`to` refer to that
 * region's doc. The plain `range_rects_json` export stays body-only.
 */
export function range_rects_region_json(display_list: string, region: string, part_id: string, from: number, to: number): string;

/**
 * Register a font for measurement from raw sfnt bytes; returns the font id
 * that `measure_paragraph_json` inputs reference in their `fontChains`.
 * Malformed bytes (attacker-controlled embedded fonts) are rejected as an
 * error at this boundary, mirroring `FontStore::register`.
 */
export function register_measure_font(bytes: Uint8Array): number;

/**
 * Rezip from a JS object `{ [path]: Uint8Array }` into a DOCX byte array.
 */
export function rezip_docx(entries: any): Uint8Array;

export function sanitizeOoxml(data: Uint8Array, expected_format: string): Uint8Array;

/**
 * Serializes an S10 request.
 */
export function serialize_docx_s10(request_json: string): string;

/**
 * Serializes an S11 request.
 */
export function serialize_docx_s11(request_json: string): string;

/**
 * Serializes an S12 request.
 */
export function serialize_docx_s12(request_json: string): string;

/**
 * Unzip a DOCX; returns a JS object `{ [path]: Uint8Array }`.
 */
export function unzip_docx(data: Uint8Array): any;

/**
 * wasm wrapper over [`session::update_display_list`]: apply a page-delta
 * update to a stored display list so an incremental rebuild re-parses only
 * its changed pages. `Err` closes the handle first, so the caller's fallback
 * (a fresh [`open_display_list`]) can never race a half-updated list.
 */
export function update_display_list(handle: number, update: string): void;

export function vertical_move_by_handle(handle: number, position: number, direction: string, goal_x: number): string;

export function vertical_move_json(display_list: string, position: number, direction: string, goal_x: number): string;

/**
 * Writes a DOCX from a typed model and original package.
 */
export function write_docx_s13_wasm(request_json: string, original_docx: Uint8Array): Uint8Array;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly __wbg_editsession_free: (a: number, b: number) => void;
    readonly editsession_accept_change: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_add_comment: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number) => [number, number, number, number];
    readonly editsession_apply_delete: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly editsession_apply_delete_profiled: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly editsession_apply_input: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly editsession_apply_input_profile_json: (a: number) => [number, number];
    readonly editsession_apply_input_profiled: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly editsession_apply_local_update: (a: number, b: number, c: number) => [number, number];
    readonly editsession_apply_paragraph_style: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number, j: number, k: number, l: number, m: number, n: number, o: number) => [number, number];
    readonly editsession_apply_raw_ops: (a: number, b: number, c: number, d: number, e: number) => [number, number];
    readonly editsession_apply_seed_raw_ops: (a: number, b: number, c: number, d: number, e: number) => [number, number];
    readonly editsession_apply_update: (a: number, b: number, c: number) => [number, number];
    readonly editsession_apply_update_with_inference: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_build_display_list_frame: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly editsession_build_display_list_json: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_can_redo: (a: number) => number;
    readonly editsession_can_undo: (a: number) => number;
    readonly editsession_cell_selection: (a: number) => [number, number, number, number];
    readonly editsession_clear_content_control_value: (a: number, b: number, c: number) => [number, number];
    readonly editsession_clear_formatting: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number) => [number, number];
    readonly editsession_clear_measure_fonts: (a: number) => void;
    readonly editsession_clear_update_event_observation: (a: number) => void;
    readonly editsession_clear_update_observer: (a: number) => void;
    readonly editsession_client_id: (a: number) => number;
    readonly editsession_create_story: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number) => [number, number, number, number];
    readonly editsession_delete_column: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_delete_range: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number, j: number, k: number, l: number, m: number) => [number, number, number, number];
    readonly editsession_delete_row: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number, number];
    readonly editsession_delete_story: (a: number, b: number, c: number) => [number, number];
    readonly editsession_delete_table: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_display_hit_test_regions_json: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly editsession_display_range_rects_json: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_display_range_rects_region_json: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number, number];
    readonly editsession_display_vertical_move_json: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly editsession_drain_update_event: (a: number) => [number, number];
    readonly editsession_encode_diff: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_encode_state: (a: number) => [number, number];
    readonly editsession_encode_state_vector: (a: number) => [number, number];
    readonly editsession_encode_sticky_position: (a: number, b: number, c: number, d: number, e: number, f: number) => [number, number, number, number];
    readonly editsession_encoded_selection: (a: number) => [number, number, number, number];
    readonly editsession_format_range: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number, j: number, k: number) => [number, number];
    readonly editsession_insert_column: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly editsession_insert_image: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number, j: number, k: number, l: number) => [number, number, number, number];
    readonly editsession_insert_page_break: (a: number, b: number, c: number, d: number, e: number, f: number) => [number, number];
    readonly editsession_insert_row: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number) => [number, number, number, number];
    readonly editsession_insert_section_break: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number) => [number, number];
    readonly editsession_insert_table: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number, j: number, k: number, l: number) => [number, number, number, number];
    readonly editsession_insert_text: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number, j: number, k: number, l: number) => [number, number, number, number];
    readonly editsession_insert_watermark: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number) => [number, number];
    readonly editsession_layout_document_json: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_layout_document_with_regions_json: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_layout_font_requirements_json: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_list_revisions: (a: number) => [number, number, number, number];
    readonly editsession_load_json: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_locate_paragraph: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly editsession_materialize_docx: (a: number) => [number, number, number, number];
    readonly editsession_measure_paragraph_json: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_merge_cells: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_merge_paragraphs: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number) => [number, number, number, number];
    readonly editsession_new: (a: number) => [number, number, number];
    readonly editsession_open_docx: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly editsession_outline_glyph_json: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_paragraph_spans: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_paragraphs: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_redo: (a: number) => number;
    readonly editsession_redo_depth: (a: number) => number;
    readonly editsession_register_measure_font: (a: number, b: number, c: number) => [number, number, number];
    readonly editsession_reject_change: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_replace_range: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number, j: number, k: number, l: number, m: number, n: number, o: number) => [number, number, number, number];
    readonly editsession_resident_caret_snapshot_json: (a: number) => [number, number, number, number];
    readonly editsession_resolve_comment: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_resolve_encoded_selection: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number, number];
    readonly editsession_resolve_sticky_position: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly editsession_seed_from_docx: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_selection: (a: number) => [number, number, number, number];
    readonly editsession_selection_context: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number) => [number, number, number, number];
    readonly editsession_set_cell_borders: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly editsession_set_cell_selection: (a: number, b: number, c: number) => [number, number];
    readonly editsession_set_cell_shading: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly editsession_set_cell_text_format: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly editsession_set_column_width: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly editsession_set_content_control_value: (a: number, b: number, c: number, d: number, e: number) => [number, number];
    readonly editsession_set_content_control_value_at: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number) => [number, number];
    readonly editsession_set_hyperlink: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number, j: number, k: number) => [number, number];
    readonly editsession_set_image_geometry: (a: number, b: number, c: number, d: number, e: number) => [number, number];
    readonly editsession_set_paragraph_attr: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number];
    readonly editsession_set_paragraph_attrs: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number, j: number, k: number, l: number, m: number, n: number, o: number) => [number, number];
    readonly editsession_set_selection: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number) => [number, number];
    readonly editsession_set_table_width: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly editsession_set_update_observer: (a: number, b: any) => [number, number];
    readonly editsession_split_cell: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly editsession_split_paragraph: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number, j: number) => [number, number, number, number];
    readonly editsession_start_update_event_observation: (a: number) => [number, number];
    readonly editsession_story_checksum: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_story_ids: (a: number) => [number, number];
    readonly editsession_story_len: (a: number, b: number, c: number) => [number, number, number];
    readonly editsession_story_segments: (a: number, b: number, c: number) => [number, number, number, number];
    readonly editsession_toggle_mark: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number, j: number, k: number) => [number, number];
    readonly editsession_track_table_undo: (a: number, b: number, c: number) => [number, number];
    readonly editsession_track_undo: (a: number, b: number, c: number) => [number, number];
    readonly editsession_undo: (a: number) => number;
    readonly editsession_undo_depth: (a: number) => number;
    readonly editsession_yrs_blocks_for_story: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly editsession_load: (a: number, b: number, c: number) => [number, number];
    readonly parse_docx_relationships: (a: number, b: number) => [number, number, number, number];
    readonly parse_docx_s2: (a: number, b: number) => [number, number, number, number];
    readonly parse_docx_s3: (a: number, b: number) => [number, number, number, number];
    readonly parse_docx_s4: (a: number, b: number) => [number, number, number, number];
    readonly parse_docx_s5: (a: number, b: number) => [number, number, number, number];
    readonly parse_docx_s6: (a: number, b: number) => [number, number, number, number];
    readonly parse_docx_s7: (a: number, b: number) => [number, number, number, number];
    readonly parse_docx_s8: (a: number, b: number) => [number, number, number, number];
    readonly parse_docx_s9: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly parse_relationships_xml: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly serialize_docx_s10: (a: number, b: number) => [number, number, number, number];
    readonly serialize_docx_s11: (a: number, b: number) => [number, number, number, number];
    readonly serialize_docx_s12: (a: number, b: number) => [number, number, number, number];
    readonly write_docx_s13_wasm: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly rezip_docx: (a: any) => [number, number, number, number];
    readonly sanitizeOoxml: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly unzip_docx: (a: number, b: number) => [number, number, number];
    readonly build_display_list_json: (a: number, b: number) => [number, number, number, number];
    readonly hit_test_json: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly hit_test_regions_by_handle: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly hit_test_regions_json: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly layout_document_json: (a: number, b: number) => [number, number, number, number];
    readonly measure_paragraph_json: (a: number, b: number) => [number, number, number, number];
    readonly open_display_list: (a: number, b: number) => [number, number, number];
    readonly outline_glyph_json: (a: number, b: number) => [number, number, number, number];
    readonly range_rects_by_handle: (a: number, b: number, c: number) => [number, number, number, number];
    readonly range_rects_json: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly range_rects_region_by_handle: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number, number];
    readonly range_rects_region_json: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number) => [number, number, number, number];
    readonly register_measure_font: (a: number, b: number) => [number, number, number];
    readonly update_display_list: (a: number, b: number, c: number) => [number, number];
    readonly vertical_move_by_handle: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly vertical_move_json: (a: number, b: number, c: number, d: number, e: number, f: number) => [number, number, number, number];
    readonly install_panic_hook: () => void;
    readonly close_display_list: (a: number) => void;
    readonly clear_measure_fonts: () => void;
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
    readonly __wbindgen_exn_store: (a: number) => void;
    readonly __externref_table_alloc: () => number;
    readonly __wbindgen_externrefs: WebAssembly.Table;
    readonly __externref_table_dealloc: (a: number) => void;
    readonly __wbindgen_free: (a: number, b: number, c: number) => void;
    readonly __externref_drop_slice: (a: number, b: number) => void;
    readonly __wbindgen_start: () => void;
}

export type SyncInitInput = BufferSource | WebAssembly.Module;

/**
 * Instantiates the given `module`, which can either be bytes or
 * a precompiled `WebAssembly.Module`.
 *
 * @param {{ module: SyncInitInput }} module - Passing `SyncInitInput` directly is deprecated.
 *
 * @returns {InitOutput}
 */
export function initSync(module: { module: SyncInitInput } | SyncInitInput): InitOutput;

/**
 * If `module_or_path` is {RequestInfo} or {URL}, makes a request and
 * for everything else, calls `WebAssembly.instantiate` directly.
 *
 * @param {{ module_or_path: InitInput | Promise<InitInput> }} module_or_path - Passing `InitInput` directly is deprecated.
 *
 * @returns {Promise<InitOutput>}
 */
export default function __wbg_init (module_or_path?: { module_or_path: InitInput | Promise<InitInput> } | InitInput | Promise<InitInput>): Promise<InitOutput>;
