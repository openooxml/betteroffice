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
     * Accepts tracked changes (S4b): pending insertions become plain content,
     * pending deletions are carried out; `pPrIns` marks clear (the split
     * stays), `pPrDel` marks join with the following paragraph (its pPr
     * survives). `target_json`: `{"revisionId": string}` for one coalesced
     * revision (any story) or
     * `{"story","startPara","startOffset","endPara","endOffset"}` for a Loc
     * range. Receipt: `{"revisionIds": [string, …]}` — the revision ids
     * resolved. Resolving never stamps a new revision.
     */
    accept_change(target_json: string): string;
    /**
     * Adds a sticky-anchored comment. `ranges_json`:
     * `[{"story","startPara","startOffset","endPara","endOffset"}, …]`;
     * `body_json` is any JSON value. Receipt: `{"commentId"}`.
     */
    add_comment(ranges_json: string, author: string, date: string, body_json: string): string;
    /**
     * Apply one ordinary collapsed character deletion (or adjacent paragraph
     * merge at a boundary) and return the resulting resident FrameDelta.
     */
    apply_delete(direction: string, expected_frame_epoch: number): Uint8Array;
    /**
     * Instrumented twin of `apply_delete`, used only by opt-in browser perf
     * traces. Keeping this separate leaves the production hot path timer-free.
     */
    apply_delete_profiled(direction: string, expected_frame_epoch: number): Uint8Array;
    /**
     * Apply one ordinary collapsed body-text insertion and return the
     * resulting FrameDelta. Selection, measurement inputs, pagination
     * checkpoints, and display state all remain resident in this session.
     */
    apply_input(text: string, expected_frame_epoch: number): Uint8Array;
    /**
     * Last opt-in `apply_input_profiled` stage timings as a compact JSON object.
     */
    apply_input_profile_json(): string;
    /**
     * Instrumented twin of `apply_input`, used only by opt-in browser perf
     * traces. Keeping this separate leaves the production hot path timer-free.
     */
    apply_input_profiled(text: string, expected_frame_epoch: number): Uint8Array;
    /**
     * Applies an update produced by this document's dedicated local worker.
     * The local origin lets the main replica's UndoManager retain ownership of
     * the edit; remote/collaboration updates must use `apply_update` instead.
     */
    apply_local_update(update: Uint8Array): void;
    /**
     * Applies a paragraph style id to every paragraph intersecting
     * `[start, end)`. With no host style resolver at this boundary, this is
     * the PM fallback path: write `pStyle` without fabricating a resolved
     * paragraph/run formatting projection.
     */
    apply_paragraph_style(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number, style_id: string, author_name?: string | null, author_date?: string | null): void;
    /**
     * Applies a batch of raw story mutations in ONE transaction — the
     * coexistence bridge's mirror-into-yrs path (not a user-intent op). `ops_json`
     * is `[{ "op":"insert"|"delete"|"format"|"insertEmbed"|"setEmbedAttr"
     * |"setComment"|"removeComment", "index", … }, …]`; each op's index (and each
     * `setComment` `[start, end)` range) is read against the story state after all
     * prior ops in the batch. Attributes/payloads are faithful mirrors of the
     * bridge's lowered PM state (tracked-change stamps arrive inside `attrs`;
     * comments are keyed by the PM comment id and anchored sticky, side-map only).
     */
    apply_raw_ops(story: string, ops_json: string): void;
    /**
     * Applies a remote/incremental yrs v1 update.
     */
    apply_update(update: Uint8Array): void;
    apply_update_with_inference(update: Uint8Array): string;
    /**
     * Binary FrameDelta v1 display output. The returned `Vec<u8>` is exposed
     * by wasm-bindgen as a transferable-friendly `Uint8Array`.
     */
    build_display_list_frame(input: string, expected_frame_epoch: number): Uint8Array;
    /**
     * Build display primitives against the same resident font store used by
     * this session's measurement path.
     */
    build_display_list_json(input: string): string;
    can_redo(): boolean;
    can_undo(): boolean;
    /**
     * Resolves the current local cell selection, or `null` before the host
     * establishes one. Deleted endpoints clamp to a surviving nearby cell.
     */
    cell_selection(): string;
    /**
     * Clears the authored `value` from a stable-id content-control embed.
     */
    clear_content_control_value(embed_id: string): void;
    /**
     * Clears every direct formatting attribute over `[start, end)`, while
     * retaining hyperlinks and tracked-change stamps.
     */
    clear_formatting(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number): void;
    /**
     * Clear this editing wasm's resident measurement fonts.
     */
    clear_measure_fonts(): void;
    clear_update_event_observation(): void;
    /**
     * Drops the update observer registered by [`EditSession::set_update_observer`].
     */
    clear_update_observer(): void;
    client_id(): number;
    /**
     * Adds a story with one paragraph. Receipt: `{"paraId"}` (the final
     * pilcrow's paragraph).
     */
    create_story(story_id: string, initial_text: string, p_style: string, alignment: string): string;
    /**
     * Deletes every column covered by an explicit cell range.
     */
    delete_column(range_json: string): string;
    /**
     * Deletes `[start, end)` given as two Locs in one story. A range whose
     * ends sit in different paragraphs spans the boundary pilcrows, so the
     * plain delete also merges (the pilcrow-as-character dividend).
     * Suggesting mode retains the content with a `del` revision instead.
     * Receipt: `{"revisionId": string|null}`.
     */
    delete_range(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number, author_name?: string | null, author_date?: string | null): string;
    /**
     * Deletes every row covered by an explicit cell range.
     */
    delete_row(range_json: string, author_name?: string | null, author_date?: string | null): string;
    /**
     * Removes one complete story (used for unreachable table-cell stories).
     */
    delete_story(story_id: string): void;
    /**
     * Removes one complete table plus all of its reachable cell stories.
     */
    delete_table(table_json: string): string;
    /**
     * Hit-test without re-serializing the resident display list through JS.
     */
    display_hit_test_regions_json(page_index: number, x: number, y: number): string;
    /**
     * Body range geometry without a display-list JSON round trip.
     */
    display_range_rects_json(from: number, to: number): string;
    /**
     * Region-scoped range geometry without a display-list JSON round trip.
     */
    display_range_rects_region_json(region: string, r_id: string, from: number, to: number): string;
    display_vertical_move_json(position: number, direction: string, goal_x: number): string;
    drain_update_event(): Uint8Array;
    encode_diff(remote_state_vector: Uint8Array): Uint8Array;
    /**
     * Full document state as one yrs v1 update (Yjs wire format).
     */
    encode_state(): Uint8Array;
    encode_state_vector(): Uint8Array;
    /**
     * Encodes one paragraph location as a sticky position.
     */
    encode_sticky_position(story: string, para_id: string, offset: number): Uint8Array;
    encoded_selection(): string;
    /**
     * Applies a set-valued, tri-state inline formatting delta over
     * `[start, end)` in one transaction. Omitted fields are kept and `null`
     * fields are cleared.
     */
    format_range(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number, delta_json: string): void;
    /**
     * Inserts a column left (`after = false`) or right (`after = true`) of a cell.
     */
    insert_column(at_json: string, after: boolean): string;
    /**
     * Inserts one native inline image embed at a paragraph-keyed location.
     */
    insert_image(story: string, para_id: string, offset: number, payload_json: string, author_name?: string | null, author_date?: string | null): string;
    /**
     * Inserts a native page-break embed at a Loc.
     */
    insert_page_break(story: string, para_id: string, offset: number): void;
    /**
     * Inserts a row above (`after = false`) or below (`after = true`) a cell.
     */
    insert_row(at_json: string, after: boolean, author_name?: string | null, author_date?: string | null): string;
    /**
     * Inserts a native section-break embed at a Loc.
     */
    insert_section_break(story: string, para_id: string, offset: number, break_type: string): void;
    /**
     * Inserts a rectangular structural table at a paragraph-keyed location.
     */
    insert_table(story: string, para_id: string, offset: number, rows: number, columns: number, author_name?: string | null, author_date?: string | null): string;
    /**
     * Inserts paragraph-break-free text at `(story, para_id, offset)`.
     * Receipt: `{"revisionId": string|null}` (non-null in suggesting mode).
     */
    insert_text(story: string, para_id: string, offset: number, text: string, author_name?: string | null, author_date?: string | null): string;
    /**
     * Inserts a typed watermark embed at a Loc.
     */
    insert_watermark(story: string, para_id: string, offset: number, watermark_json: string): void;
    /**
     * Paginate and retain the measured input and Layout. The full JSON return
     * remains the migration parity bridge until binary frames consume it.
     */
    layout_document_json(input: string): string;
    /**
     * Paginate and compose section/page regions inside the resident engine.
     */
    layout_document_with_regions_json(input: string): string;
    layout_font_requirements_json(input: string): string;
    /**
     * Every tracked-change run/paragraph-mark revision across all stories,
     * in deterministic story/position order.
     */
    list_revisions(): string;
    /**
     * Hydrates this replica from an encoded yrs update (the bytes form of
     * `load` — typically another replica's `encode_state()` output).
     */
    load(update: Uint8Array): void;
    /**
     * Seeds stories from JSON (the json form of `load`):
     * `[{"storyId","paragraphs":[{"text","pStyle"?,"alignment"?}, …]}, …]`.
     * Paragraph text must not contain paragraph breaks. Returns
     * `{storyId: [paraId, …]}` in document order. This is an S1 seeding
     * scaffold composed from public ops; the real `load(ParsedDocument)`
     * belongs to the ops track.
     */
    load_json(stories_json: string): string;
    /**
     * `{"start","end"}` — the paragraph's story span; `end` is its pilcrow's
     * index, so `end - start` is the paragraph length (`offset` domain).
     */
    locate_paragraph(story: string, para_id: string): string;
    /**
     * Materializes the retained canonical package for compatibility APIs.
     */
    materialize_docx(): string | undefined;
    /**
     * Paragraph-measure compatibility export on the resident engine module.
     */
    measure_paragraph_json(input: string): string;
    /**
     * Merges a rectangular cell range into its top-left cell.
     */
    merge_cells(range_json: string): string;
    /**
     * Merges `para_id` with the FOLLOWING paragraph by deleting (plain) or
     * `del`-marking (suggesting) its pilcrow. Errors on the story's final
     * paragraph. Receipt: `{"revisionId": string|null}`.
     */
    merge_paragraphs(story: string, para_id: string, author_name?: string | null, author_date?: string | null): string;
    /**
     * Creates a replica. `client_id` must be a non-negative safe integer —
     * the host allocates it (yjs-style random 32-bit ids are fine).
     */
    constructor(client_id: number);
    /**
     * Parses a DOCX and optionally seeds its editable stories.
     */
    open_docx(bytes: Uint8Array, seed_stories: boolean): string;
    /**
     * Resolve one glyph outline from the session's resident font store.
     */
    outline_glyph_json(font_id: number, glyph_id: number): string;
    /**
     * Compact paragraph-position projection in one story traversal:
     * `[{"paraId","length"}]`. Length counts UTF-16 text and inline embed
     * units before each paragraph's pilcrow. The JS input shim uses this
     * instead of crossing the wasm boundary once per paragraph.
     */
    paragraph_spans(story: string): string;
    /**
     * `[{"paraId","text","properties"}]` in document order. `properties`
     * carries pStyle/alignment plus any op-set extras.
     */
    paragraphs(story: string): string;
    /**
     * Reapplies the latest locally undone transaction.
     */
    redo(): boolean;
    /**
     * Current local redo stack size. Zero before a story starts tracking.
     */
    redo_depth(): number;
    /**
     * Register font bytes in this editing wasm's resident measurement store.
     * Returned ids are valid for measurement and display work in this module.
     */
    register_measure_font(bytes: Uint8Array): number;
    /**
     * Rejects tracked changes — the inverse of [`EditSession::accept_change`]:
     * pending insertions roll back, pending deletions restore their text;
     * `pPrIns` marks join back with the following paragraph, `pPrDel` marks
     * clear (the split stays). Same target and receipt shapes.
     */
    reject_change(target_json: string): string;
    /**
     * Replaces `[start, end)` with text in one transaction. The inserted text
     * adopts the first replaced unit's formatting; in suggesting mode the
     * deletion and insertion share one revision id.
     */
    replace_range(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number, text: string, author_name?: string | null, author_date?: string | null): string;
    resident_caret_snapshot_json(): string;
    /**
     * Current offsets of a comment's sticky anchors:
     * `[{"story","start","end"}]`. Errors when an anchor no longer resolves.
     */
    resolve_comment(comment_id: string): string;
    resolve_encoded_selection(story: string, anchor: Uint8Array, head: Uint8Array): string;
    /**
     * Resolves one encoded sticky position to a paragraph location.
     */
    resolve_sticky_position(story: string, position: Uint8Array): string;
    /**
     * Parses a DOCX, seeds its stories, and returns thin host metadata.
     */
    seed_from_docx(bytes: Uint8Array): string;
    /**
     * Resolves this peer's current sticky selection as two public Locs, or
     * `null` before the host establishes an initial selection.
     */
    selection(): string;
    /**
     * Aggregated toolbar/a11y state over one paragraph-addressed story
     * range. Toggle marks are `true`, `false`, or `"mixed"`; value marks
     * are their uniform value or `null` when absent/mixed.
     */
    selection_context(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number): string;
    /**
     * Replaces the selected cells' complete border property object.
     */
    set_cell_borders(range_json: string, borders_json: string): string;
    /**
     * Stores a rectangular anchor-cell → head-cell selection outside the yrs
     * document. `range_json` is a [`TableRange`]. The table embed is held by a
     * sticky index and the endpoints by stable cell-story identity.
     */
    set_cell_selection(range_json: string): void;
    /**
     * Sets/clears selected cells' background color (hex without or with `#`).
     */
    set_cell_shading(range_json: string, color?: string | null): string;
    /**
     * Merges a JSON cell-format patch into every selected cell's `tcPr`.
     */
    set_cell_text_format(range_json: string, patch_json: string): string;
    /**
     * Sets one grid-column width in twips.
     */
    set_column_width(at_json: string, width_twips: number): string;
    /**
     * Sets the authored `value` on a stable-id content-control embed.
     */
    set_content_control_value(embed_id: string, value_json: string): void;
    /**
     * Sets the authored `value` on a content-control embed at a paragraph-keyed
     * position. This is the fallback for valid controls that have no authored
     * `w:id`/tag and therefore cannot be addressed by stable payload identity.
     */
    set_content_control_value_at(story: string, para_id: string, offset: number, value_json: string): void;
    /**
     * Sets or clears the protected hyperlink attribute over `[start, end)`.
     * `hyperlink_json` is an object (`{href, tooltip?, rId?}`) or `null`.
     */
    set_hyperlink(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number, hyperlink_json: string): void;
    /**
     * Commits image geometry fields to a stable-id image embed in one
     * transaction. `null` fields clear; `other` is flattened into the payload.
     */
    set_image_geometry(embed_id: string, geometry_json: string): void;
    /**
     * Sets one paragraph property (any JSON value) on `para_id`'s pilcrow.
     * `paraId` / the embed discriminator are reserved.
     */
    set_paragraph_attr(para_id: string, key: string, value_json: string): void;
    /**
     * Applies a tri-state paragraph-property delta to every paragraph
     * intersecting `[start, end)` in one transaction.
     */
    set_paragraph_attrs(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number, attrs_json: string, author_name?: string | null, author_date?: string | null): void;
    /**
     * Stores this peer's anchor/head as sticky positions. `Assoc::After`
     * makes a collapsed caret advance with text inserted at the caret.
     */
    set_selection(story: string, anchor_para: string, anchor_offset: number, head_para: string, head_offset: number): void;
    /**
     * Sets the table-wide preferred width in twips.
     */
    set_table_width(table_json: string, width_twips: number): string;
    /**
     * Subscribes `callback(update: Uint8Array)` to every committed
     * transaction (v1 encoding — feed it straight to `apply_update` on a
     * peer). One observer per session; a second call replaces the first.
     * The facade fans out to multiple JS listeners over this single hook.
     */
    set_update_observer(callback: Function): void;
    /**
     * Splits the cell covering `at` into the requested grid. Omitted
     * dimensions unmerge the cell into its existing covered slots.
     */
    split_cell(at_json: string, rows?: number | null, columns?: number | null): string;
    /**
     * Splits a paragraph at `(story, para_id, offset)` by inserting one
     * pilcrow. Under the S1 split the FIRST half keeps the original paraId and
     * the SECOND half is re-minted. Receipt:
     * `{"firstParaId","secondParaId","revisionId": string|null}`.
     */
    split_paragraph(story: string, para_id: string, offset: number, author_name?: string | null, author_date?: string | null): string;
    start_update_event_observation(): void;
    /**
     * The story's `canonical-stream-v1` FNV-1a checksum as a decimal string
     * (u64 exceeds JS safe-integer range). The coexistence watchdog compares
     * this against the PM projector's checksum after every mirrored edit.
     */
    story_checksum(story: string): string;
    /**
     * Story ids currently in the document, sorted for determinism.
     */
    story_ids(): string[];
    /**
     * Story length in UTF-16 units (every embed, pilcrows included, = 1).
     */
    story_len(story: string): number;
    /**
     * The raw formatted-segment view (the render bridge's input):
     * `[{"kind":"text","text",…} | {"kind":"pilcrow","paraId","properties",…}
     * | {"kind":"embed",…}]`, each with `"attributes"` (run marks plus
     * `ins`/`del` revision values).
     */
    story_segments(story: string): string;
    /**
     * Applies one run mark over `[start, end)` (two Locs in one story). Simple
     * marks toggle; font/size/color set (see [`apply_mark`]). `mark_json`:
     * `{"type":"bold"|"italic"|"underline"|"strike"|"superscript"|"subscript"} |
     * {"type":"fontFamily"|"color","value":string} |
     * {"type":"fontSize","value":number}`.
     */
    toggle_mark(story: string, start_para: string, start_offset: number, end_para: string, end_offset: number, mark_json: string): void;
    /**
     * Starts local undo tracking for a structural table transaction. Besides
     * the parent story (which owns the table embed), the stories root must be
     * in scope so undo/redo also removes/restores cell-story map entries.
     */
    track_table_undo(story: string): void;
    /**
     * Starts local-origin undo tracking for one story. Hosts call this lazily
     * after import/seeding but before the first direct input operation, so the
     * initial document is not an undo step.
     */
    track_undo(story: string): void;
    /**
     * Reverts the latest local-origin transaction. Remote/system mirror
     * transactions are excluded by `DocUndoManager`'s tracked-origin policy.
     */
    undo(): boolean;
    /**
     * Current local undo stack size. Zero before a story starts tracking.
     */
    undo_depth(): number;
    /**
     * Lowers a story through the resident Rust bridge. Errors with an
     * unsupported-embed message on any non-native content until that class is
     * promoted to native.
     * `env_json` carries theme colors, the default tab stop, and list numeric
     * ids (see [`parse_render_env`]).
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
 * point in, PM position (or `null`) as JSON out.
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
 * `{"region":"body"|"header"|"footer","rId"?,"pos":n|null}` (or `"null"` for
 * an out-of-range page). The legacy `hit_test_json` export stays body-only.
 */
export function hit_test_regions_json(display_list: string, page_index: number, x: number, y: number): string;

/**
 * wasm wrapper over [`layout_to_json`].
 */
export function layout_document_json(input: string): string;

/**
 * wasm wrapper over [`ooxml_text::measure_paragraph_json`]: measurement
 * input JSON in, `ParagraphExtent` JSON out. An `Err` whose message starts
 * with `"UNSUPPORTED"` means the caller must fall back to browser
 * measurement for that block.
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
 * wasm wrapper over [`session::range_rects_by_handle`]: range rects against a
 * stored display list. `Err` on an unknown/closed handle so the caller can
 * fall back to [`range_rects_json`].
 */
export function range_rects_by_handle(handle: number, from: number, to: number): string;

/**
 * wasm wrapper over [`hit::range_rects_json`]: display-list JSON + PM range
 * in, JSON array of page-local rects out.
 */
export function range_rects_json(display_list: string, from: number, to: number): string;

/**
 * wasm wrapper over [`session::range_rects_region_by_handle`]: region-aware
 * range rects against a stored display list. `region` is
 * `"body" | "header" | "footer"`; `r_id` scopes header/footer to one HF part.
 * `Err` on an unknown/closed handle so the caller can fall back to
 * [`range_rects_region_json`].
 */
export function range_rects_region_by_handle(handle: number, region: string, r_id: string, from: number, to: number): string;

/**
 * wasm wrapper over [`hit::range_rects_region_json`]: region-aware range rects.
 * `region` is `"body" | "header" | "footer"`; `r_id` scopes a header/footer to
 * one HF part (empty for body / match-any). The `from`/`to` refer to that
 * region's PM doc. The legacy `range_rects_json` export stays body-only.
 */
export function range_rects_region_json(display_list: string, region: string, r_id: string, from: number, to: number): string;

/**
 * Register a font for measurement from raw sfnt bytes; returns the font id
 * that `measure_paragraph_json` inputs reference in their `fontChains`.
 * Malformed bytes (attacker-controlled embedded fonts) are rejected as an
 * error at this boundary, mirroring `FontStore::register`.
 */
export function register_measure_font(bytes: Uint8Array): number;

/**
 * wasm wrapper over [`session::update_display_list`]: apply a page-delta
 * update to a stored display list so an incremental rebuild re-parses only
 * its changed pages. `Err` closes the handle first, so the caller's fallback
 * (a fresh [`open_display_list`]) can never race a half-updated list.
 */
export function update_display_list(handle: number, update: string): void;

export function vertical_move_by_handle(handle: number, position: number, direction: string, goal_x: number): string;

export function vertical_move_json(display_list: string, position: number, direction: string, goal_x: number): string;

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
    readonly close_display_list: (a: number) => void;
    readonly clear_measure_fonts: () => void;
    readonly __wbindgen_exn_store: (a: number) => void;
    readonly __externref_table_alloc: () => number;
    readonly __wbindgen_externrefs: WebAssembly.Table;
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
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
