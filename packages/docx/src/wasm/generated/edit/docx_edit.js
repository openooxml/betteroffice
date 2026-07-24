/* @ts-self-types="./docx_edit.d.ts" */

/**
 * One yrs replica of the DOCX editing model, held for a JS host.
 *
 * Owns the [`EditingDoc`] plus the (single) JS update observer. The JS facade
 * multiplexes its own listener set over that one callback.
 */
export class EditSession {
    __destroy_into_raw() {
        const ptr = this.__wbg_ptr;
        this.__wbg_ptr = 0;
        EditSessionFinalization.unregister(this);
        return ptr;
    }
    free() {
        const ptr = this.__destroy_into_raw();
        wasm.__wbg_editsession_free(ptr, 0);
    }
    /**
     * Accepts tracked changes (S4b): pending insertions become plain content,
     * pending deletions are carried out; `pPrIns` marks clear (the split
     * stays), `pPrDel` marks join with the following paragraph (its pPr
     * survives). `target_json`: `{"revisionId": string}` for one coalesced
     * revision (any story) or
     * `{"story","startPara","startOffset","endPara","endOffset"}` for a Loc
     * range. Receipt: `{"revisionIds": [string, …]}` — the revision ids
     * resolved. Resolving never stamps a new revision.
     * @param {string} target_json
     * @returns {string}
     */
    accept_change(target_json) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(target_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_accept_change(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * Adds a sticky-anchored comment. `ranges_json`:
     * `[{"story","startPara","startOffset","endPara","endOffset"}, …]`;
     * `body_json` is any JSON value. Receipt: `{"commentId"}`.
     * @param {string} ranges_json
     * @param {string} author
     * @param {string} date
     * @param {string} body_json
     * @returns {string}
     */
    add_comment(ranges_json, author, date, body_json) {
        let deferred6_0;
        let deferred6_1;
        try {
            const ptr0 = passStringToWasm0(ranges_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(author, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            const ptr2 = passStringToWasm0(date, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len2 = WASM_VECTOR_LEN;
            const ptr3 = passStringToWasm0(body_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len3 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_add_comment(this.__wbg_ptr, ptr0, len0, ptr1, len1, ptr2, len2, ptr3, len3);
            var ptr5 = ret[0];
            var len5 = ret[1];
            if (ret[3]) {
                ptr5 = 0; len5 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred6_0 = ptr5;
            deferred6_1 = len5;
            return getStringFromWasm0(ptr5, len5);
        } finally {
            wasm.__wbindgen_free(deferred6_0, deferred6_1, 1);
        }
    }
    /**
     * Apply one ordinary collapsed character deletion (or adjacent paragraph
     * merge at a boundary) and return the resulting resident FrameDelta.
     * @param {string} direction
     * @param {number} expected_frame_epoch
     * @returns {Uint8Array}
     */
    apply_delete(direction, expected_frame_epoch) {
        const ptr0 = passStringToWasm0(direction, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_apply_delete(this.__wbg_ptr, ptr0, len0, expected_frame_epoch);
        if (ret[3]) {
            throw takeFromExternrefTable0(ret[2]);
        }
        var v2 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
        wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
        return v2;
    }
    /**
     * Instrumented twin of `apply_delete`, used only by opt-in browser perf
     * traces. Keeping this separate leaves the production hot path timer-free.
     * @param {string} direction
     * @param {number} expected_frame_epoch
     * @returns {Uint8Array}
     */
    apply_delete_profiled(direction, expected_frame_epoch) {
        const ptr0 = passStringToWasm0(direction, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_apply_delete_profiled(this.__wbg_ptr, ptr0, len0, expected_frame_epoch);
        if (ret[3]) {
            throw takeFromExternrefTable0(ret[2]);
        }
        var v2 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
        wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
        return v2;
    }
    /**
     * Apply one ordinary collapsed body-text insertion and return the
     * resulting FrameDelta. Selection, measurement inputs, pagination
     * checkpoints, and display state all remain resident in this session.
     * @param {string} text
     * @param {number} expected_frame_epoch
     * @returns {Uint8Array}
     */
    apply_input(text, expected_frame_epoch) {
        const ptr0 = passStringToWasm0(text, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_apply_input(this.__wbg_ptr, ptr0, len0, expected_frame_epoch);
        if (ret[3]) {
            throw takeFromExternrefTable0(ret[2]);
        }
        var v2 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
        wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
        return v2;
    }
    /**
     * Last opt-in `apply_input_profiled` stage timings as a compact JSON object.
     * @returns {string}
     */
    apply_input_profile_json() {
        let deferred1_0;
        let deferred1_1;
        try {
            const ret = wasm.editsession_apply_input_profile_json(this.__wbg_ptr);
            deferred1_0 = ret[0];
            deferred1_1 = ret[1];
            return getStringFromWasm0(ret[0], ret[1]);
        } finally {
            wasm.__wbindgen_free(deferred1_0, deferred1_1, 1);
        }
    }
    /**
     * Instrumented twin of `apply_input`, used only by opt-in browser perf
     * traces. Keeping this separate leaves the production hot path timer-free.
     * @param {string} text
     * @param {number} expected_frame_epoch
     * @returns {Uint8Array}
     */
    apply_input_profiled(text, expected_frame_epoch) {
        const ptr0 = passStringToWasm0(text, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_apply_input_profiled(this.__wbg_ptr, ptr0, len0, expected_frame_epoch);
        if (ret[3]) {
            throw takeFromExternrefTable0(ret[2]);
        }
        var v2 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
        wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
        return v2;
    }
    /**
     * Applies an update produced by this document's dedicated local worker.
     * The local origin lets the main replica's UndoManager retain ownership of
     * the edit; remote/collaboration updates must use `apply_update` instead.
     * @param {Uint8Array} update
     */
    apply_local_update(update) {
        const ptr0 = passArray8ToWasm0(update, wasm.__wbindgen_malloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_apply_local_update(this.__wbg_ptr, ptr0, len0);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Applies a paragraph style id to every paragraph intersecting
     * `[start, end)`. With no host style resolver at this boundary, this is
     * the PM fallback path: write `pStyle` without fabricating a resolved
     * paragraph/run formatting projection.
     * @param {string} story
     * @param {string} start_para
     * @param {number} start_offset
     * @param {string} end_para
     * @param {number} end_offset
     * @param {string} style_id
     * @param {string | null} [author_name]
     * @param {string | null} [author_date]
     */
    apply_paragraph_style(story, start_para, start_offset, end_para, end_offset, style_id, author_name, author_date) {
        const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(start_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ptr2 = passStringToWasm0(end_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len2 = WASM_VECTOR_LEN;
        const ptr3 = passStringToWasm0(style_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len3 = WASM_VECTOR_LEN;
        var ptr4 = isLikeNone(author_name) ? 0 : passStringToWasm0(author_name, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        var len4 = WASM_VECTOR_LEN;
        var ptr5 = isLikeNone(author_date) ? 0 : passStringToWasm0(author_date, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        var len5 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_apply_paragraph_style(this.__wbg_ptr, ptr0, len0, ptr1, len1, start_offset, ptr2, len2, end_offset, ptr3, len3, ptr4, len4, ptr5, len5);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Applies a batch of raw story mutations in ONE transaction — the
     * coexistence bridge's mirror-into-yrs path (not a user-intent op). `ops_json`
     * is `[{ "op":"insert"|"delete"|"format"|"insertEmbed"|"setEmbedAttr"
     * |"setComment"|"removeComment", "index", … }, …]`; each op's index (and each
     * `setComment` `[start, end)` range) is read against the story state after all
     * prior ops in the batch. Attributes/payloads are faithful mirrors of the
     * bridge's lowered PM state (tracked-change stamps arrive inside `attrs`;
     * comments are keyed by the PM comment id and anchored sticky, side-map only).
     * @param {string} story
     * @param {string} ops_json
     */
    apply_raw_ops(story, ops_json) {
        const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(ops_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_apply_raw_ops(this.__wbg_ptr, ptr0, len0, ptr1, len1);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Applies a remote/incremental yrs v1 update.
     * @param {Uint8Array} update
     */
    apply_update(update) {
        const ptr0 = passArray8ToWasm0(update, wasm.__wbindgen_malloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_apply_update(this.__wbg_ptr, ptr0, len0);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Binary FrameDelta v1 display output. The returned `Vec<u8>` is exposed
     * by wasm-bindgen as a transferable-friendly `Uint8Array`.
     * @param {string} input
     * @param {number} expected_frame_epoch
     * @returns {Uint8Array}
     */
    build_display_list_frame(input, expected_frame_epoch) {
        const ptr0 = passStringToWasm0(input, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_build_display_list_frame(this.__wbg_ptr, ptr0, len0, expected_frame_epoch);
        if (ret[3]) {
            throw takeFromExternrefTable0(ret[2]);
        }
        var v2 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
        wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
        return v2;
    }
    /**
     * Build display primitives against the same resident font store used by
     * this session's measurement path.
     * @param {string} input
     * @returns {string}
     */
    build_display_list_json(input) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(input, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_build_display_list_json(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * @returns {boolean}
     */
    can_redo() {
        const ret = wasm.editsession_can_redo(this.__wbg_ptr);
        return ret !== 0;
    }
    /**
     * @returns {boolean}
     */
    can_undo() {
        const ret = wasm.editsession_can_undo(this.__wbg_ptr);
        return ret !== 0;
    }
    /**
     * Resolves the current local cell selection, or `null` before the host
     * establishes one. Deleted endpoints clamp to a surviving nearby cell.
     * @returns {string}
     */
    cell_selection() {
        let deferred2_0;
        let deferred2_1;
        try {
            const ret = wasm.editsession_cell_selection(this.__wbg_ptr);
            var ptr1 = ret[0];
            var len1 = ret[1];
            if (ret[3]) {
                ptr1 = 0; len1 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred2_0 = ptr1;
            deferred2_1 = len1;
            return getStringFromWasm0(ptr1, len1);
        } finally {
            wasm.__wbindgen_free(deferred2_0, deferred2_1, 1);
        }
    }
    /**
     * Clears the authored `value` from a stable-id content-control embed.
     * @param {string} embed_id
     */
    clear_content_control_value(embed_id) {
        const ptr0 = passStringToWasm0(embed_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_clear_content_control_value(this.__wbg_ptr, ptr0, len0);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Clears every direct formatting attribute over `[start, end)`, while
     * retaining hyperlinks and tracked-change stamps.
     * @param {string} story
     * @param {string} start_para
     * @param {number} start_offset
     * @param {string} end_para
     * @param {number} end_offset
     */
    clear_formatting(story, start_para, start_offset, end_para, end_offset) {
        const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(start_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ptr2 = passStringToWasm0(end_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len2 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_clear_formatting(this.__wbg_ptr, ptr0, len0, ptr1, len1, start_offset, ptr2, len2, end_offset);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Clear this editing wasm's resident measurement fonts.
     */
    clear_measure_fonts() {
        wasm.editsession_clear_measure_fonts(this.__wbg_ptr);
    }
    clear_update_event_observation() {
        wasm.editsession_clear_update_event_observation(this.__wbg_ptr);
    }
    /**
     * Drops the update observer registered by [`EditSession::set_update_observer`].
     */
    clear_update_observer() {
        wasm.editsession_clear_update_observer(this.__wbg_ptr);
    }
    /**
     * @returns {number}
     */
    client_id() {
        const ret = wasm.editsession_client_id(this.__wbg_ptr);
        return ret;
    }
    /**
     * Adds a story with one paragraph. Receipt: `{"paraId"}` (the final
     * pilcrow's paragraph).
     * @param {string} story_id
     * @param {string} initial_text
     * @param {string} p_style
     * @param {string} alignment
     * @returns {string}
     */
    create_story(story_id, initial_text, p_style, alignment) {
        let deferred6_0;
        let deferred6_1;
        try {
            const ptr0 = passStringToWasm0(story_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(initial_text, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            const ptr2 = passStringToWasm0(p_style, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len2 = WASM_VECTOR_LEN;
            const ptr3 = passStringToWasm0(alignment, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len3 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_create_story(this.__wbg_ptr, ptr0, len0, ptr1, len1, ptr2, len2, ptr3, len3);
            var ptr5 = ret[0];
            var len5 = ret[1];
            if (ret[3]) {
                ptr5 = 0; len5 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred6_0 = ptr5;
            deferred6_1 = len5;
            return getStringFromWasm0(ptr5, len5);
        } finally {
            wasm.__wbindgen_free(deferred6_0, deferred6_1, 1);
        }
    }
    /**
     * Deletes every column covered by an explicit cell range.
     * @param {string} range_json
     * @returns {string}
     */
    delete_column(range_json) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(range_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_delete_column(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * Deletes `[start, end)` given as two Locs in one story. A range whose
     * ends sit in different paragraphs spans the boundary pilcrows, so the
     * plain delete also merges (the pilcrow-as-character dividend).
     * Suggesting mode retains the content with a `del` revision instead.
     * Receipt: `{"revisionId": string|null}`.
     * @param {string} story
     * @param {string} start_para
     * @param {number} start_offset
     * @param {string} end_para
     * @param {number} end_offset
     * @param {string | null} [author_name]
     * @param {string | null} [author_date]
     * @returns {string}
     */
    delete_range(story, start_para, start_offset, end_para, end_offset, author_name, author_date) {
        let deferred7_0;
        let deferred7_1;
        try {
            const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(start_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            const ptr2 = passStringToWasm0(end_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len2 = WASM_VECTOR_LEN;
            var ptr3 = isLikeNone(author_name) ? 0 : passStringToWasm0(author_name, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len3 = WASM_VECTOR_LEN;
            var ptr4 = isLikeNone(author_date) ? 0 : passStringToWasm0(author_date, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len4 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_delete_range(this.__wbg_ptr, ptr0, len0, ptr1, len1, start_offset, ptr2, len2, end_offset, ptr3, len3, ptr4, len4);
            var ptr6 = ret[0];
            var len6 = ret[1];
            if (ret[3]) {
                ptr6 = 0; len6 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred7_0 = ptr6;
            deferred7_1 = len6;
            return getStringFromWasm0(ptr6, len6);
        } finally {
            wasm.__wbindgen_free(deferred7_0, deferred7_1, 1);
        }
    }
    /**
     * Deletes every row covered by an explicit cell range.
     * @param {string} range_json
     * @param {string | null} [author_name]
     * @param {string | null} [author_date]
     * @returns {string}
     */
    delete_row(range_json, author_name, author_date) {
        let deferred5_0;
        let deferred5_1;
        try {
            const ptr0 = passStringToWasm0(range_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            var ptr1 = isLikeNone(author_name) ? 0 : passStringToWasm0(author_name, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len1 = WASM_VECTOR_LEN;
            var ptr2 = isLikeNone(author_date) ? 0 : passStringToWasm0(author_date, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len2 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_delete_row(this.__wbg_ptr, ptr0, len0, ptr1, len1, ptr2, len2);
            var ptr4 = ret[0];
            var len4 = ret[1];
            if (ret[3]) {
                ptr4 = 0; len4 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred5_0 = ptr4;
            deferred5_1 = len4;
            return getStringFromWasm0(ptr4, len4);
        } finally {
            wasm.__wbindgen_free(deferred5_0, deferred5_1, 1);
        }
    }
    /**
     * Removes one complete story (used for unreachable table-cell stories).
     * @param {string} story_id
     */
    delete_story(story_id) {
        const ptr0 = passStringToWasm0(story_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_delete_story(this.__wbg_ptr, ptr0, len0);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Removes one complete table plus all of its reachable cell stories.
     * @param {string} table_json
     * @returns {string}
     */
    delete_table(table_json) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(table_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_delete_table(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * Hit-test without re-serializing the resident display list through JS.
     * @param {number} page_index
     * @param {number} x
     * @param {number} y
     * @returns {string}
     */
    display_hit_test_regions_json(page_index, x, y) {
        let deferred2_0;
        let deferred2_1;
        try {
            const ret = wasm.editsession_display_hit_test_regions_json(this.__wbg_ptr, page_index, x, y);
            var ptr1 = ret[0];
            var len1 = ret[1];
            if (ret[3]) {
                ptr1 = 0; len1 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred2_0 = ptr1;
            deferred2_1 = len1;
            return getStringFromWasm0(ptr1, len1);
        } finally {
            wasm.__wbindgen_free(deferred2_0, deferred2_1, 1);
        }
    }
    /**
     * Body range geometry without a display-list JSON round trip.
     * @param {number} from
     * @param {number} to
     * @returns {string}
     */
    display_range_rects_json(from, to) {
        let deferred2_0;
        let deferred2_1;
        try {
            const ret = wasm.editsession_display_range_rects_json(this.__wbg_ptr, from, to);
            var ptr1 = ret[0];
            var len1 = ret[1];
            if (ret[3]) {
                ptr1 = 0; len1 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred2_0 = ptr1;
            deferred2_1 = len1;
            return getStringFromWasm0(ptr1, len1);
        } finally {
            wasm.__wbindgen_free(deferred2_0, deferred2_1, 1);
        }
    }
    /**
     * Region-scoped range geometry without a display-list JSON round trip.
     * @param {string} region
     * @param {string} r_id
     * @param {number} from
     * @param {number} to
     * @returns {string}
     */
    display_range_rects_region_json(region, r_id, from, to) {
        let deferred4_0;
        let deferred4_1;
        try {
            const ptr0 = passStringToWasm0(region, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(r_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_display_range_rects_region_json(this.__wbg_ptr, ptr0, len0, ptr1, len1, from, to);
            var ptr3 = ret[0];
            var len3 = ret[1];
            if (ret[3]) {
                ptr3 = 0; len3 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred4_0 = ptr3;
            deferred4_1 = len3;
            return getStringFromWasm0(ptr3, len3);
        } finally {
            wasm.__wbindgen_free(deferred4_0, deferred4_1, 1);
        }
    }
    /**
     * @returns {Uint8Array}
     */
    drain_update_event() {
        const ret = wasm.editsession_drain_update_event(this.__wbg_ptr);
        var v1 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
        wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
        return v1;
    }
    /**
     * @param {Uint8Array} remote_state_vector
     * @returns {Uint8Array}
     */
    encode_diff(remote_state_vector) {
        const ptr0 = passArray8ToWasm0(remote_state_vector, wasm.__wbindgen_malloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_encode_diff(this.__wbg_ptr, ptr0, len0);
        if (ret[3]) {
            throw takeFromExternrefTable0(ret[2]);
        }
        var v2 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
        wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
        return v2;
    }
    /**
     * Full document state as one yrs v1 update (Yjs wire format).
     * @returns {Uint8Array}
     */
    encode_state() {
        const ret = wasm.editsession_encode_state(this.__wbg_ptr);
        var v1 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
        wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
        return v1;
    }
    /**
     * @returns {Uint8Array}
     */
    encode_state_vector() {
        const ret = wasm.editsession_encode_state_vector(this.__wbg_ptr);
        var v1 = getArrayU8FromWasm0(ret[0], ret[1]).slice();
        wasm.__wbindgen_free(ret[0], ret[1] * 1, 1);
        return v1;
    }
    /**
     * Applies a set-valued, tri-state inline formatting delta over
     * `[start, end)` in one transaction. Omitted fields are kept and `null`
     * fields are cleared.
     * @param {string} story
     * @param {string} start_para
     * @param {number} start_offset
     * @param {string} end_para
     * @param {number} end_offset
     * @param {string} delta_json
     */
    format_range(story, start_para, start_offset, end_para, end_offset, delta_json) {
        const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(start_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ptr2 = passStringToWasm0(end_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len2 = WASM_VECTOR_LEN;
        const ptr3 = passStringToWasm0(delta_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len3 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_format_range(this.__wbg_ptr, ptr0, len0, ptr1, len1, start_offset, ptr2, len2, end_offset, ptr3, len3);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Inserts a column left (`after = false`) or right (`after = true`) of a cell.
     * @param {string} at_json
     * @param {boolean} after
     * @returns {string}
     */
    insert_column(at_json, after) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(at_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_insert_column(this.__wbg_ptr, ptr0, len0, after);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * Inserts one native inline image embed at a paragraph-keyed location.
     * @param {string} story
     * @param {string} para_id
     * @param {number} offset
     * @param {string} payload_json
     * @param {string | null} [author_name]
     * @param {string | null} [author_date]
     * @returns {string}
     */
    insert_image(story, para_id, offset, payload_json, author_name, author_date) {
        let deferred7_0;
        let deferred7_1;
        try {
            const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(para_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            const ptr2 = passStringToWasm0(payload_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len2 = WASM_VECTOR_LEN;
            var ptr3 = isLikeNone(author_name) ? 0 : passStringToWasm0(author_name, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len3 = WASM_VECTOR_LEN;
            var ptr4 = isLikeNone(author_date) ? 0 : passStringToWasm0(author_date, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len4 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_insert_image(this.__wbg_ptr, ptr0, len0, ptr1, len1, offset, ptr2, len2, ptr3, len3, ptr4, len4);
            var ptr6 = ret[0];
            var len6 = ret[1];
            if (ret[3]) {
                ptr6 = 0; len6 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred7_0 = ptr6;
            deferred7_1 = len6;
            return getStringFromWasm0(ptr6, len6);
        } finally {
            wasm.__wbindgen_free(deferred7_0, deferred7_1, 1);
        }
    }
    /**
     * Inserts a native page-break embed at a Loc.
     * @param {string} story
     * @param {string} para_id
     * @param {number} offset
     */
    insert_page_break(story, para_id, offset) {
        const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(para_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_insert_page_break(this.__wbg_ptr, ptr0, len0, ptr1, len1, offset);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Inserts a row above (`after = false`) or below (`after = true`) a cell.
     * @param {string} at_json
     * @param {boolean} after
     * @param {string | null} [author_name]
     * @param {string | null} [author_date]
     * @returns {string}
     */
    insert_row(at_json, after, author_name, author_date) {
        let deferred5_0;
        let deferred5_1;
        try {
            const ptr0 = passStringToWasm0(at_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            var ptr1 = isLikeNone(author_name) ? 0 : passStringToWasm0(author_name, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len1 = WASM_VECTOR_LEN;
            var ptr2 = isLikeNone(author_date) ? 0 : passStringToWasm0(author_date, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len2 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_insert_row(this.__wbg_ptr, ptr0, len0, after, ptr1, len1, ptr2, len2);
            var ptr4 = ret[0];
            var len4 = ret[1];
            if (ret[3]) {
                ptr4 = 0; len4 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred5_0 = ptr4;
            deferred5_1 = len4;
            return getStringFromWasm0(ptr4, len4);
        } finally {
            wasm.__wbindgen_free(deferred5_0, deferred5_1, 1);
        }
    }
    /**
     * Inserts a native section-break embed at a Loc.
     * @param {string} story
     * @param {string} para_id
     * @param {number} offset
     * @param {string} break_type
     */
    insert_section_break(story, para_id, offset, break_type) {
        const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(para_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ptr2 = passStringToWasm0(break_type, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len2 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_insert_section_break(this.__wbg_ptr, ptr0, len0, ptr1, len1, offset, ptr2, len2);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Inserts a rectangular structural table at a paragraph-keyed location.
     * @param {string} story
     * @param {string} para_id
     * @param {number} offset
     * @param {number} rows
     * @param {number} columns
     * @param {string | null} [author_name]
     * @param {string | null} [author_date]
     * @returns {string}
     */
    insert_table(story, para_id, offset, rows, columns, author_name, author_date) {
        let deferred6_0;
        let deferred6_1;
        try {
            const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(para_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            var ptr2 = isLikeNone(author_name) ? 0 : passStringToWasm0(author_name, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len2 = WASM_VECTOR_LEN;
            var ptr3 = isLikeNone(author_date) ? 0 : passStringToWasm0(author_date, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len3 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_insert_table(this.__wbg_ptr, ptr0, len0, ptr1, len1, offset, rows, columns, ptr2, len2, ptr3, len3);
            var ptr5 = ret[0];
            var len5 = ret[1];
            if (ret[3]) {
                ptr5 = 0; len5 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred6_0 = ptr5;
            deferred6_1 = len5;
            return getStringFromWasm0(ptr5, len5);
        } finally {
            wasm.__wbindgen_free(deferred6_0, deferred6_1, 1);
        }
    }
    /**
     * Inserts paragraph-break-free text at `(story, para_id, offset)`.
     * Receipt: `{"revisionId": string|null}` (non-null in suggesting mode).
     * @param {string} story
     * @param {string} para_id
     * @param {number} offset
     * @param {string} text
     * @param {string | null} [author_name]
     * @param {string | null} [author_date]
     * @returns {string}
     */
    insert_text(story, para_id, offset, text, author_name, author_date) {
        let deferred7_0;
        let deferred7_1;
        try {
            const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(para_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            const ptr2 = passStringToWasm0(text, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len2 = WASM_VECTOR_LEN;
            var ptr3 = isLikeNone(author_name) ? 0 : passStringToWasm0(author_name, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len3 = WASM_VECTOR_LEN;
            var ptr4 = isLikeNone(author_date) ? 0 : passStringToWasm0(author_date, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len4 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_insert_text(this.__wbg_ptr, ptr0, len0, ptr1, len1, offset, ptr2, len2, ptr3, len3, ptr4, len4);
            var ptr6 = ret[0];
            var len6 = ret[1];
            if (ret[3]) {
                ptr6 = 0; len6 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred7_0 = ptr6;
            deferred7_1 = len6;
            return getStringFromWasm0(ptr6, len6);
        } finally {
            wasm.__wbindgen_free(deferred7_0, deferred7_1, 1);
        }
    }
    /**
     * Inserts a typed watermark embed at a Loc.
     * @param {string} story
     * @param {string} para_id
     * @param {number} offset
     * @param {string} watermark_json
     */
    insert_watermark(story, para_id, offset, watermark_json) {
        const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(para_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ptr2 = passStringToWasm0(watermark_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len2 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_insert_watermark(this.__wbg_ptr, ptr0, len0, ptr1, len1, offset, ptr2, len2);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Paginate and retain the measured input and Layout. The full JSON return
     * remains the migration parity bridge until binary frames consume it.
     * @param {string} input
     * @returns {string}
     */
    layout_document_json(input) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(input, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_layout_document_json(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * Paginate and compose section/page regions inside the resident engine.
     * @param {string} input
     * @returns {string}
     */
    layout_document_with_regions_json(input) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(input, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_layout_document_with_regions_json(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * @param {string} input
     * @returns {string}
     */
    layout_font_requirements_json(input) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(input, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_layout_font_requirements_json(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * Every tracked-change run/paragraph-mark revision across all stories,
     * in deterministic story/position order.
     * @returns {string}
     */
    list_revisions() {
        let deferred2_0;
        let deferred2_1;
        try {
            const ret = wasm.editsession_list_revisions(this.__wbg_ptr);
            var ptr1 = ret[0];
            var len1 = ret[1];
            if (ret[3]) {
                ptr1 = 0; len1 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred2_0 = ptr1;
            deferred2_1 = len1;
            return getStringFromWasm0(ptr1, len1);
        } finally {
            wasm.__wbindgen_free(deferred2_0, deferred2_1, 1);
        }
    }
    /**
     * Hydrates this replica from an encoded yrs update (the bytes form of
     * `load` — typically another replica's `encode_state()` output).
     * @param {Uint8Array} update
     */
    load(update) {
        const ptr0 = passArray8ToWasm0(update, wasm.__wbindgen_malloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_load(this.__wbg_ptr, ptr0, len0);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Seeds stories from JSON (the json form of `load`):
     * `[{"storyId","paragraphs":[{"text","pStyle"?,"alignment"?}, …]}, …]`.
     * Paragraph text must not contain paragraph breaks. Returns
     * `{storyId: [paraId, …]}` in document order. This is an S1 seeding
     * scaffold composed from public ops; the real `load(ParsedDocument)`
     * belongs to the ops track.
     * @param {string} stories_json
     * @returns {string}
     */
    load_json(stories_json) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(stories_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_load_json(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * `{"start","end"}` — the paragraph's story span; `end` is its pilcrow's
     * index, so `end - start` is the paragraph length (`offset` domain).
     * @param {string} story
     * @param {string} para_id
     * @returns {string}
     */
    locate_paragraph(story, para_id) {
        let deferred4_0;
        let deferred4_1;
        try {
            const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(para_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_locate_paragraph(this.__wbg_ptr, ptr0, len0, ptr1, len1);
            var ptr3 = ret[0];
            var len3 = ret[1];
            if (ret[3]) {
                ptr3 = 0; len3 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred4_0 = ptr3;
            deferred4_1 = len3;
            return getStringFromWasm0(ptr3, len3);
        } finally {
            wasm.__wbindgen_free(deferred4_0, deferred4_1, 1);
        }
    }
    /**
     * Paragraph-measure compatibility export on the resident engine module.
     * @param {string} input
     * @returns {string}
     */
    measure_paragraph_json(input) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(input, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_measure_paragraph_json(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * Merges a rectangular cell range into its top-left cell.
     * @param {string} range_json
     * @returns {string}
     */
    merge_cells(range_json) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(range_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_merge_cells(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * Merges `para_id` with the FOLLOWING paragraph by deleting (plain) or
     * `del`-marking (suggesting) its pilcrow. Errors on the story's final
     * paragraph. Receipt: `{"revisionId": string|null}`.
     * @param {string} story
     * @param {string} para_id
     * @param {string | null} [author_name]
     * @param {string | null} [author_date]
     * @returns {string}
     */
    merge_paragraphs(story, para_id, author_name, author_date) {
        let deferred6_0;
        let deferred6_1;
        try {
            const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(para_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            var ptr2 = isLikeNone(author_name) ? 0 : passStringToWasm0(author_name, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len2 = WASM_VECTOR_LEN;
            var ptr3 = isLikeNone(author_date) ? 0 : passStringToWasm0(author_date, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len3 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_merge_paragraphs(this.__wbg_ptr, ptr0, len0, ptr1, len1, ptr2, len2, ptr3, len3);
            var ptr5 = ret[0];
            var len5 = ret[1];
            if (ret[3]) {
                ptr5 = 0; len5 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred6_0 = ptr5;
            deferred6_1 = len5;
            return getStringFromWasm0(ptr5, len5);
        } finally {
            wasm.__wbindgen_free(deferred6_0, deferred6_1, 1);
        }
    }
    /**
     * Creates a replica. `client_id` must be a non-negative safe integer —
     * the host allocates it (yjs-style random 32-bit ids are fine).
     * @param {number} client_id
     */
    constructor(client_id) {
        const ret = wasm.editsession_new(client_id);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        this.__wbg_ptr = ret[0];
        EditSessionFinalization.register(this, this.__wbg_ptr, this);
        return this;
    }
    /**
     * Resolve one glyph outline from the session's resident font store.
     * @param {number} font_id
     * @param {number} glyph_id
     * @returns {string}
     */
    outline_glyph_json(font_id, glyph_id) {
        let deferred2_0;
        let deferred2_1;
        try {
            const ret = wasm.editsession_outline_glyph_json(this.__wbg_ptr, font_id, glyph_id);
            var ptr1 = ret[0];
            var len1 = ret[1];
            if (ret[3]) {
                ptr1 = 0; len1 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred2_0 = ptr1;
            deferred2_1 = len1;
            return getStringFromWasm0(ptr1, len1);
        } finally {
            wasm.__wbindgen_free(deferred2_0, deferred2_1, 1);
        }
    }
    /**
     * Compact paragraph-position projection in one story traversal:
     * `[{"paraId","length"}]`. Length counts UTF-16 text and inline embed
     * units before each paragraph's pilcrow. The JS input shim uses this
     * instead of crossing the wasm boundary once per paragraph.
     * @param {string} story
     * @returns {string}
     */
    paragraph_spans(story) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_paragraph_spans(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * `[{"paraId","text","properties"}]` in document order. `properties`
     * carries pStyle/alignment plus any op-set extras.
     * @param {string} story
     * @returns {string}
     */
    paragraphs(story) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_paragraphs(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * Reapplies the latest locally undone transaction.
     * @returns {boolean}
     */
    redo() {
        const ret = wasm.editsession_redo(this.__wbg_ptr);
        return ret !== 0;
    }
    /**
     * Current local redo stack size. Zero before a story starts tracking.
     * @returns {number}
     */
    redo_depth() {
        const ret = wasm.editsession_redo_depth(this.__wbg_ptr);
        return ret >>> 0;
    }
    /**
     * Register font bytes in this editing wasm's resident measurement store.
     * Returned ids are valid for measurement and display work in this module.
     * @param {Uint8Array} bytes
     * @returns {number}
     */
    register_measure_font(bytes) {
        const ptr0 = passArray8ToWasm0(bytes, wasm.__wbindgen_malloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_register_measure_font(this.__wbg_ptr, ptr0, len0);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return ret[0] >>> 0;
    }
    /**
     * Rejects tracked changes — the inverse of [`EditSession::accept_change`]:
     * pending insertions roll back, pending deletions restore their text;
     * `pPrIns` marks join back with the following paragraph, `pPrDel` marks
     * clear (the split stays). Same target and receipt shapes.
     * @param {string} target_json
     * @returns {string}
     */
    reject_change(target_json) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(target_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_reject_change(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * Replaces `[start, end)` with text in one transaction. The inserted text
     * adopts the first replaced unit's formatting; in suggesting mode the
     * deletion and insertion share one revision id.
     * @param {string} story
     * @param {string} start_para
     * @param {number} start_offset
     * @param {string} end_para
     * @param {number} end_offset
     * @param {string} text
     * @param {string | null} [author_name]
     * @param {string | null} [author_date]
     * @returns {string}
     */
    replace_range(story, start_para, start_offset, end_para, end_offset, text, author_name, author_date) {
        let deferred8_0;
        let deferred8_1;
        try {
            const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(start_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            const ptr2 = passStringToWasm0(end_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len2 = WASM_VECTOR_LEN;
            const ptr3 = passStringToWasm0(text, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len3 = WASM_VECTOR_LEN;
            var ptr4 = isLikeNone(author_name) ? 0 : passStringToWasm0(author_name, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len4 = WASM_VECTOR_LEN;
            var ptr5 = isLikeNone(author_date) ? 0 : passStringToWasm0(author_date, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len5 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_replace_range(this.__wbg_ptr, ptr0, len0, ptr1, len1, start_offset, ptr2, len2, end_offset, ptr3, len3, ptr4, len4, ptr5, len5);
            var ptr7 = ret[0];
            var len7 = ret[1];
            if (ret[3]) {
                ptr7 = 0; len7 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred8_0 = ptr7;
            deferred8_1 = len7;
            return getStringFromWasm0(ptr7, len7);
        } finally {
            wasm.__wbindgen_free(deferred8_0, deferred8_1, 1);
        }
    }
    /**
     * @returns {string}
     */
    resident_caret_snapshot_json() {
        let deferred2_0;
        let deferred2_1;
        try {
            const ret = wasm.editsession_resident_caret_snapshot_json(this.__wbg_ptr);
            var ptr1 = ret[0];
            var len1 = ret[1];
            if (ret[3]) {
                ptr1 = 0; len1 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred2_0 = ptr1;
            deferred2_1 = len1;
            return getStringFromWasm0(ptr1, len1);
        } finally {
            wasm.__wbindgen_free(deferred2_0, deferred2_1, 1);
        }
    }
    /**
     * Current offsets of a comment's sticky anchors:
     * `[{"story","start","end"}]`. Errors when an anchor no longer resolves.
     * @param {string} comment_id
     * @returns {string}
     */
    resolve_comment(comment_id) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(comment_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_resolve_comment(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * Resolves this peer's current sticky selection as two public Locs, or
     * `null` before the host establishes an initial selection.
     * @returns {string}
     */
    selection() {
        let deferred2_0;
        let deferred2_1;
        try {
            const ret = wasm.editsession_selection(this.__wbg_ptr);
            var ptr1 = ret[0];
            var len1 = ret[1];
            if (ret[3]) {
                ptr1 = 0; len1 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred2_0 = ptr1;
            deferred2_1 = len1;
            return getStringFromWasm0(ptr1, len1);
        } finally {
            wasm.__wbindgen_free(deferred2_0, deferred2_1, 1);
        }
    }
    /**
     * Aggregated toolbar/a11y state over one paragraph-addressed story
     * range. Toggle marks are `true`, `false`, or `"mixed"`; value marks
     * are their uniform value or `null` when absent/mixed.
     * @param {string} story
     * @param {string} start_para
     * @param {number} start_offset
     * @param {string} end_para
     * @param {number} end_offset
     * @returns {string}
     */
    selection_context(story, start_para, start_offset, end_para, end_offset) {
        let deferred5_0;
        let deferred5_1;
        try {
            const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(start_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            const ptr2 = passStringToWasm0(end_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len2 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_selection_context(this.__wbg_ptr, ptr0, len0, ptr1, len1, start_offset, ptr2, len2, end_offset);
            var ptr4 = ret[0];
            var len4 = ret[1];
            if (ret[3]) {
                ptr4 = 0; len4 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred5_0 = ptr4;
            deferred5_1 = len4;
            return getStringFromWasm0(ptr4, len4);
        } finally {
            wasm.__wbindgen_free(deferred5_0, deferred5_1, 1);
        }
    }
    /**
     * Replaces the selected cells' complete border property object.
     * @param {string} range_json
     * @param {string} borders_json
     * @returns {string}
     */
    set_cell_borders(range_json, borders_json) {
        let deferred4_0;
        let deferred4_1;
        try {
            const ptr0 = passStringToWasm0(range_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(borders_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_set_cell_borders(this.__wbg_ptr, ptr0, len0, ptr1, len1);
            var ptr3 = ret[0];
            var len3 = ret[1];
            if (ret[3]) {
                ptr3 = 0; len3 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred4_0 = ptr3;
            deferred4_1 = len3;
            return getStringFromWasm0(ptr3, len3);
        } finally {
            wasm.__wbindgen_free(deferred4_0, deferred4_1, 1);
        }
    }
    /**
     * Stores a rectangular anchor-cell → head-cell selection outside the yrs
     * document. `range_json` is a [`TableRange`]. The table embed is held by a
     * sticky index and the endpoints by stable cell-story identity.
     * @param {string} range_json
     */
    set_cell_selection(range_json) {
        const ptr0 = passStringToWasm0(range_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_set_cell_selection(this.__wbg_ptr, ptr0, len0);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Sets/clears selected cells' background color (hex without or with `#`).
     * @param {string} range_json
     * @param {string | null} [color]
     * @returns {string}
     */
    set_cell_shading(range_json, color) {
        let deferred4_0;
        let deferred4_1;
        try {
            const ptr0 = passStringToWasm0(range_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            var ptr1 = isLikeNone(color) ? 0 : passStringToWasm0(color, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len1 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_set_cell_shading(this.__wbg_ptr, ptr0, len0, ptr1, len1);
            var ptr3 = ret[0];
            var len3 = ret[1];
            if (ret[3]) {
                ptr3 = 0; len3 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred4_0 = ptr3;
            deferred4_1 = len3;
            return getStringFromWasm0(ptr3, len3);
        } finally {
            wasm.__wbindgen_free(deferred4_0, deferred4_1, 1);
        }
    }
    /**
     * Merges a JSON cell-format patch into every selected cell's `tcPr`.
     * @param {string} range_json
     * @param {string} patch_json
     * @returns {string}
     */
    set_cell_text_format(range_json, patch_json) {
        let deferred4_0;
        let deferred4_1;
        try {
            const ptr0 = passStringToWasm0(range_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(patch_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_set_cell_text_format(this.__wbg_ptr, ptr0, len0, ptr1, len1);
            var ptr3 = ret[0];
            var len3 = ret[1];
            if (ret[3]) {
                ptr3 = 0; len3 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred4_0 = ptr3;
            deferred4_1 = len3;
            return getStringFromWasm0(ptr3, len3);
        } finally {
            wasm.__wbindgen_free(deferred4_0, deferred4_1, 1);
        }
    }
    /**
     * Sets one grid-column width in twips.
     * @param {string} at_json
     * @param {number} width_twips
     * @returns {string}
     */
    set_column_width(at_json, width_twips) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(at_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_set_column_width(this.__wbg_ptr, ptr0, len0, width_twips);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * Sets the authored `value` on a stable-id content-control embed.
     * @param {string} embed_id
     * @param {string} value_json
     */
    set_content_control_value(embed_id, value_json) {
        const ptr0 = passStringToWasm0(embed_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(value_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_set_content_control_value(this.__wbg_ptr, ptr0, len0, ptr1, len1);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Sets the authored `value` on a content-control embed at a paragraph-keyed
     * position. This is the fallback for valid controls that have no authored
     * `w:id`/tag and therefore cannot be addressed by stable payload identity.
     * @param {string} story
     * @param {string} para_id
     * @param {number} offset
     * @param {string} value_json
     */
    set_content_control_value_at(story, para_id, offset, value_json) {
        const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(para_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ptr2 = passStringToWasm0(value_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len2 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_set_content_control_value_at(this.__wbg_ptr, ptr0, len0, ptr1, len1, offset, ptr2, len2);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Sets or clears the protected hyperlink attribute over `[start, end)`.
     * `hyperlink_json` is an object (`{href, tooltip?, rId?}`) or `null`.
     * @param {string} story
     * @param {string} start_para
     * @param {number} start_offset
     * @param {string} end_para
     * @param {number} end_offset
     * @param {string} hyperlink_json
     */
    set_hyperlink(story, start_para, start_offset, end_para, end_offset, hyperlink_json) {
        const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(start_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ptr2 = passStringToWasm0(end_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len2 = WASM_VECTOR_LEN;
        const ptr3 = passStringToWasm0(hyperlink_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len3 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_set_hyperlink(this.__wbg_ptr, ptr0, len0, ptr1, len1, start_offset, ptr2, len2, end_offset, ptr3, len3);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Commits image geometry fields to a stable-id image embed in one
     * transaction. `null` fields clear; `other` is flattened into the payload.
     * @param {string} embed_id
     * @param {string} geometry_json
     */
    set_image_geometry(embed_id, geometry_json) {
        const ptr0 = passStringToWasm0(embed_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(geometry_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_set_image_geometry(this.__wbg_ptr, ptr0, len0, ptr1, len1);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Sets one paragraph property (any JSON value) on `para_id`'s pilcrow.
     * `paraId` / the embed discriminator are reserved.
     * @param {string} para_id
     * @param {string} key
     * @param {string} value_json
     */
    set_paragraph_attr(para_id, key, value_json) {
        const ptr0 = passStringToWasm0(para_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(key, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ptr2 = passStringToWasm0(value_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len2 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_set_paragraph_attr(this.__wbg_ptr, ptr0, len0, ptr1, len1, ptr2, len2);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Applies a tri-state paragraph-property delta to every paragraph
     * intersecting `[start, end)` in one transaction.
     * @param {string} story
     * @param {string} start_para
     * @param {number} start_offset
     * @param {string} end_para
     * @param {number} end_offset
     * @param {string} attrs_json
     * @param {string | null} [author_name]
     * @param {string | null} [author_date]
     */
    set_paragraph_attrs(story, start_para, start_offset, end_para, end_offset, attrs_json, author_name, author_date) {
        const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(start_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ptr2 = passStringToWasm0(end_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len2 = WASM_VECTOR_LEN;
        const ptr3 = passStringToWasm0(attrs_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len3 = WASM_VECTOR_LEN;
        var ptr4 = isLikeNone(author_name) ? 0 : passStringToWasm0(author_name, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        var len4 = WASM_VECTOR_LEN;
        var ptr5 = isLikeNone(author_date) ? 0 : passStringToWasm0(author_date, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        var len5 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_set_paragraph_attrs(this.__wbg_ptr, ptr0, len0, ptr1, len1, start_offset, ptr2, len2, end_offset, ptr3, len3, ptr4, len4, ptr5, len5);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Stores this peer's anchor/head as sticky positions. `Assoc::After`
     * makes a collapsed caret advance with text inserted at the caret.
     * @param {string} story
     * @param {string} anchor_para
     * @param {number} anchor_offset
     * @param {string} head_para
     * @param {number} head_offset
     */
    set_selection(story, anchor_para, anchor_offset, head_para, head_offset) {
        const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(anchor_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ptr2 = passStringToWasm0(head_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len2 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_set_selection(this.__wbg_ptr, ptr0, len0, ptr1, len1, anchor_offset, ptr2, len2, head_offset);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Sets the table-wide preferred width in twips.
     * @param {string} table_json
     * @param {number} width_twips
     * @returns {string}
     */
    set_table_width(table_json, width_twips) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(table_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_set_table_width(this.__wbg_ptr, ptr0, len0, width_twips);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * Subscribes `callback(update: Uint8Array)` to every committed
     * transaction (v1 encoding — feed it straight to `apply_update` on a
     * peer). One observer per session; a second call replaces the first.
     * The facade fans out to multiple JS listeners over this single hook.
     * @param {Function} callback
     */
    set_update_observer(callback) {
        const ret = wasm.editsession_set_update_observer(this.__wbg_ptr, callback);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Splits the cell covering `at` into the requested grid. Omitted
     * dimensions unmerge the cell into its existing covered slots.
     * @param {string} at_json
     * @param {number | null} [rows]
     * @param {number | null} [columns]
     * @returns {string}
     */
    split_cell(at_json, rows, columns) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(at_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_split_cell(this.__wbg_ptr, ptr0, len0, isLikeNone(rows) ? Number.MAX_SAFE_INTEGER : (rows) >>> 0, isLikeNone(columns) ? Number.MAX_SAFE_INTEGER : (columns) >>> 0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * Splits a paragraph at `(story, para_id, offset)` by inserting one
     * pilcrow. Under the S1 split the FIRST half keeps the original paraId and
     * the SECOND half is re-minted. Receipt:
     * `{"firstParaId","secondParaId","revisionId": string|null}`.
     * @param {string} story
     * @param {string} para_id
     * @param {number} offset
     * @param {string | null} [author_name]
     * @param {string | null} [author_date]
     * @returns {string}
     */
    split_paragraph(story, para_id, offset, author_name, author_date) {
        let deferred6_0;
        let deferred6_1;
        try {
            const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(para_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            var ptr2 = isLikeNone(author_name) ? 0 : passStringToWasm0(author_name, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len2 = WASM_VECTOR_LEN;
            var ptr3 = isLikeNone(author_date) ? 0 : passStringToWasm0(author_date, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            var len3 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_split_paragraph(this.__wbg_ptr, ptr0, len0, ptr1, len1, offset, ptr2, len2, ptr3, len3);
            var ptr5 = ret[0];
            var len5 = ret[1];
            if (ret[3]) {
                ptr5 = 0; len5 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred6_0 = ptr5;
            deferred6_1 = len5;
            return getStringFromWasm0(ptr5, len5);
        } finally {
            wasm.__wbindgen_free(deferred6_0, deferred6_1, 1);
        }
    }
    start_update_event_observation() {
        const ret = wasm.editsession_start_update_event_observation(this.__wbg_ptr);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * The story's `canonical-stream-v1` FNV-1a checksum as a decimal string
     * (u64 exceeds JS safe-integer range). The coexistence watchdog compares
     * this against the PM projector's checksum after every mirrored edit.
     * @param {string} story
     * @returns {string}
     */
    story_checksum(story) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_story_checksum(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * Story ids currently in the document, sorted for determinism.
     * @returns {string[]}
     */
    story_ids() {
        const ret = wasm.editsession_story_ids(this.__wbg_ptr);
        var v1 = getArrayJsValueFromWasm0(ret[0], ret[1]).slice();
        wasm.__wbindgen_free(ret[0], ret[1] * 4, 4);
        return v1;
    }
    /**
     * Story length in UTF-16 units (every embed, pilcrows included, = 1).
     * @param {string} story
     * @returns {number}
     */
    story_len(story) {
        const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_story_len(this.__wbg_ptr, ptr0, len0);
        if (ret[2]) {
            throw takeFromExternrefTable0(ret[1]);
        }
        return ret[0] >>> 0;
    }
    /**
     * The raw formatted-segment view (the render bridge's input):
     * `[{"kind":"text","text",…} | {"kind":"pilcrow","paraId","properties",…}
     * | {"kind":"embed",…}]`, each with `"attributes"` (run marks plus
     * `ins`/`del` revision values).
     * @param {string} story
     * @returns {string}
     */
    story_segments(story) {
        let deferred3_0;
        let deferred3_1;
        try {
            const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_story_segments(this.__wbg_ptr, ptr0, len0);
            var ptr2 = ret[0];
            var len2 = ret[1];
            if (ret[3]) {
                ptr2 = 0; len2 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred3_0 = ptr2;
            deferred3_1 = len2;
            return getStringFromWasm0(ptr2, len2);
        } finally {
            wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
        }
    }
    /**
     * Applies one run mark over `[start, end)` (two Locs in one story). Simple
     * marks toggle; font/size/color set (see [`apply_mark`]). `mark_json`:
     * `{"type":"bold"|"italic"|"underline"|"strike"|"superscript"|"subscript"} |
     * {"type":"fontFamily"|"color","value":string} |
     * {"type":"fontSize","value":number}`.
     * @param {string} story
     * @param {string} start_para
     * @param {number} start_offset
     * @param {string} end_para
     * @param {number} end_offset
     * @param {string} mark_json
     */
    toggle_mark(story, start_para, start_offset, end_para, end_offset, mark_json) {
        const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(start_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ptr2 = passStringToWasm0(end_para, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len2 = WASM_VECTOR_LEN;
        const ptr3 = passStringToWasm0(mark_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len3 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_toggle_mark(this.__wbg_ptr, ptr0, len0, ptr1, len1, start_offset, ptr2, len2, end_offset, ptr3, len3);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Starts local undo tracking for a structural table transaction. Besides
     * the parent story (which owns the table embed), the stories root must be
     * in scope so undo/redo also removes/restores cell-story map entries.
     * @param {string} story
     */
    track_table_undo(story) {
        const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_track_table_undo(this.__wbg_ptr, ptr0, len0);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Starts local-origin undo tracking for one story. Hosts call this lazily
     * after import/seeding but before the first direct input operation, so the
     * initial document is not an undo step.
     * @param {string} story
     */
    track_undo(story) {
        const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.editsession_track_undo(this.__wbg_ptr, ptr0, len0);
        if (ret[1]) {
            throw takeFromExternrefTable0(ret[0]);
        }
    }
    /**
     * Reverts the latest local-origin transaction. Remote/system mirror
     * transactions are excluded by `DocUndoManager`'s tracked-origin policy.
     * @returns {boolean}
     */
    undo() {
        const ret = wasm.editsession_undo(this.__wbg_ptr);
        return ret !== 0;
    }
    /**
     * Current local undo stack size. Zero before a story starts tracking.
     * @returns {number}
     */
    undo_depth() {
        const ret = wasm.editsession_undo_depth(this.__wbg_ptr);
        return ret >>> 0;
    }
    /**
     * Lowers a story through the resident Rust bridge. Errors with an
     * unsupported-embed message on any non-native content until that class is
     * promoted to native.
     * `env_json` carries theme colors, the default tab stop, and list numeric
     * ids (see [`parse_render_env`]).
     * @param {string} story
     * @param {string} env_json
     * @returns {string}
     */
    yrs_blocks_for_story(story, env_json) {
        let deferred4_0;
        let deferred4_1;
        try {
            const ptr0 = passStringToWasm0(story, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len0 = WASM_VECTOR_LEN;
            const ptr1 = passStringToWasm0(env_json, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
            const len1 = WASM_VECTOR_LEN;
            const ret = wasm.editsession_yrs_blocks_for_story(this.__wbg_ptr, ptr0, len0, ptr1, len1);
            var ptr3 = ret[0];
            var len3 = ret[1];
            if (ret[3]) {
                ptr3 = 0; len3 = 0;
                throw takeFromExternrefTable0(ret[2]);
            }
            deferred4_0 = ptr3;
            deferred4_1 = len3;
            return getStringFromWasm0(ptr3, len3);
        } finally {
            wasm.__wbindgen_free(deferred4_0, deferred4_1, 1);
        }
    }
}
if (Symbol.dispose) EditSession.prototype[Symbol.dispose] = EditSession.prototype.free;

/**
 * wasm compatibility wrapper. Resident engine users call
 * [`build_display_list_value`] and keep the typed result.
 * @param {string} input
 * @returns {string}
 */
export function build_display_list_json(input) {
    let deferred3_0;
    let deferred3_1;
    try {
        const ptr0 = passStringToWasm0(input, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.build_display_list_json(ptr0, len0);
        var ptr2 = ret[0];
        var len2 = ret[1];
        if (ret[3]) {
            ptr2 = 0; len2 = 0;
            throw takeFromExternrefTable0(ret[2]);
        }
        deferred3_0 = ptr2;
        deferred3_1 = len2;
        return getStringFromWasm0(ptr2, len2);
    } finally {
        wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
    }
}

/**
 * Drop every registered measurement font (ids restart at 0). Callers must
 * re-register before the next `measure_paragraph_json`.
 */
export function clear_measure_fonts() {
    wasm.clear_measure_fonts();
}

/**
 * wasm wrapper over [`session::close_display_list`]: drop a handle so its
 * parsed display list is freed. Idempotent.
 * @param {number} handle
 */
export function close_display_list(handle) {
    wasm.close_display_list(handle);
}

/**
 * wasm wrapper over [`hit::hit_test_json`]: display-list JSON + page-local
 * point in, PM position (or `null`) as JSON out.
 * @param {string} display_list
 * @param {number} page_index
 * @param {number} x
 * @param {number} y
 * @returns {string}
 */
export function hit_test_json(display_list, page_index, x, y) {
    let deferred3_0;
    let deferred3_1;
    try {
        const ptr0 = passStringToWasm0(display_list, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.hit_test_json(ptr0, len0, page_index, x, y);
        var ptr2 = ret[0];
        var len2 = ret[1];
        if (ret[3]) {
            ptr2 = 0; len2 = 0;
            throw takeFromExternrefTable0(ret[2]);
        }
        deferred3_0 = ptr2;
        deferred3_1 = len2;
        return getStringFromWasm0(ptr2, len2);
    } finally {
        wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
    }
}

/**
 * wasm wrapper over [`session::hit_test_regions_by_handle`]: region-aware hit
 * test against a stored display list. `Err` on an unknown/closed handle so the
 * caller can fall back to [`hit_test_regions_json`].
 * @param {number} handle
 * @param {number} page_index
 * @param {number} x
 * @param {number} y
 * @returns {string}
 */
export function hit_test_regions_by_handle(handle, page_index, x, y) {
    let deferred2_0;
    let deferred2_1;
    try {
        const ret = wasm.hit_test_regions_by_handle(handle, page_index, x, y);
        var ptr1 = ret[0];
        var len1 = ret[1];
        if (ret[3]) {
            ptr1 = 0; len1 = 0;
            throw takeFromExternrefTable0(ret[2]);
        }
        deferred2_0 = ptr1;
        deferred2_1 = len1;
        return getStringFromWasm0(ptr1, len1);
    } finally {
        wasm.__wbindgen_free(deferred2_0, deferred2_1, 1);
    }
}

/**
 * wasm wrapper over [`hit::hit_test_regions_json`]: region-aware hit test —
 * `{"region":"body"|"header"|"footer","rId"?,"pos":n|null}` (or `"null"` for
 * an out-of-range page). The legacy `hit_test_json` export stays body-only.
 * @param {string} display_list
 * @param {number} page_index
 * @param {number} x
 * @param {number} y
 * @returns {string}
 */
export function hit_test_regions_json(display_list, page_index, x, y) {
    let deferred3_0;
    let deferred3_1;
    try {
        const ptr0 = passStringToWasm0(display_list, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.hit_test_regions_json(ptr0, len0, page_index, x, y);
        var ptr2 = ret[0];
        var len2 = ret[1];
        if (ret[3]) {
            ptr2 = 0; len2 = 0;
            throw takeFromExternrefTable0(ret[2]);
        }
        deferred3_0 = ptr2;
        deferred3_1 = len2;
        return getStringFromWasm0(ptr2, len2);
    } finally {
        wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
    }
}

/**
 * wasm wrapper over [`layout_to_json`].
 * @param {string} input
 * @returns {string}
 */
export function layout_document_json(input) {
    let deferred3_0;
    let deferred3_1;
    try {
        const ptr0 = passStringToWasm0(input, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.layout_document_json(ptr0, len0);
        var ptr2 = ret[0];
        var len2 = ret[1];
        if (ret[3]) {
            ptr2 = 0; len2 = 0;
            throw takeFromExternrefTable0(ret[2]);
        }
        deferred3_0 = ptr2;
        deferred3_1 = len2;
        return getStringFromWasm0(ptr2, len2);
    } finally {
        wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
    }
}

/**
 * wasm wrapper over [`ooxml_text::measure_paragraph_json`]: measurement
 * input JSON in, `ParagraphExtent` JSON out. An `Err` whose message starts
 * with `"UNSUPPORTED"` means the caller must fall back to browser
 * measurement for that block.
 * @param {string} input
 * @returns {string}
 */
export function measure_paragraph_json(input) {
    let deferred3_0;
    let deferred3_1;
    try {
        const ptr0 = passStringToWasm0(input, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.measure_paragraph_json(ptr0, len0);
        var ptr2 = ret[0];
        var len2 = ret[1];
        if (ret[3]) {
            ptr2 = 0; len2 = 0;
            throw takeFromExternrefTable0(ret[2]);
        }
        deferred3_0 = ptr2;
        deferred3_1 = len2;
        return getStringFromWasm0(ptr2, len2);
    } finally {
        wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
    }
}

/**
 * wasm wrapper over [`session::open_display_list`]: parse a display list once
 * and return a handle the by-handle query exports reuse (no per-query
 * re-parse). The caller frees it with [`close_display_list`]. `Err` on
 * malformed JSON — the caller then stays on the JSON-arg path.
 * @param {string} display_list
 * @returns {number}
 */
export function open_display_list(display_list) {
    const ptr0 = passStringToWasm0(display_list, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.open_display_list(ptr0, len0);
    if (ret[2]) {
        throw takeFromExternrefTable0(ret[1]);
    }
    return ret[0] >>> 0;
}

/**
 * wasm wrapper over [`ooxml_text::FontStore::outline_glyph_json`]: the outline
 * of a registered font's glyph, in font design units, as JSON:
 * `{"upem":2048,"cmds":[{"t":"M","x":..,"y":..},{"t":"L","x":..,"y":..},
 * {"t":"Q","cx":..,"cy":..,"x":..,"y":..},
 * {"t":"C","c1x":..,"c1y":..,"c2x":..,"c2y":..,"x":..,"y":..},{"t":"Z"}]}`.
 * The canvas caches this per `(fontId, glyphId)` and scales by `size/upem`,
 * flipping y at draw time. `cmds` is empty for a blank glyph (space).
 * @param {number} font_id
 * @param {number} glyph_id
 * @returns {string}
 */
export function outline_glyph_json(font_id, glyph_id) {
    let deferred2_0;
    let deferred2_1;
    try {
        const ret = wasm.outline_glyph_json(font_id, glyph_id);
        var ptr1 = ret[0];
        var len1 = ret[1];
        if (ret[3]) {
            ptr1 = 0; len1 = 0;
            throw takeFromExternrefTable0(ret[2]);
        }
        deferred2_0 = ptr1;
        deferred2_1 = len1;
        return getStringFromWasm0(ptr1, len1);
    } finally {
        wasm.__wbindgen_free(deferred2_0, deferred2_1, 1);
    }
}

/**
 * wasm wrapper over [`session::range_rects_by_handle`]: range rects against a
 * stored display list. `Err` on an unknown/closed handle so the caller can
 * fall back to [`range_rects_json`].
 * @param {number} handle
 * @param {number} from
 * @param {number} to
 * @returns {string}
 */
export function range_rects_by_handle(handle, from, to) {
    let deferred2_0;
    let deferred2_1;
    try {
        const ret = wasm.range_rects_by_handle(handle, from, to);
        var ptr1 = ret[0];
        var len1 = ret[1];
        if (ret[3]) {
            ptr1 = 0; len1 = 0;
            throw takeFromExternrefTable0(ret[2]);
        }
        deferred2_0 = ptr1;
        deferred2_1 = len1;
        return getStringFromWasm0(ptr1, len1);
    } finally {
        wasm.__wbindgen_free(deferred2_0, deferred2_1, 1);
    }
}

/**
 * wasm wrapper over [`hit::range_rects_json`]: display-list JSON + PM range
 * in, JSON array of page-local rects out.
 * @param {string} display_list
 * @param {number} from
 * @param {number} to
 * @returns {string}
 */
export function range_rects_json(display_list, from, to) {
    let deferred3_0;
    let deferred3_1;
    try {
        const ptr0 = passStringToWasm0(display_list, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.range_rects_json(ptr0, len0, from, to);
        var ptr2 = ret[0];
        var len2 = ret[1];
        if (ret[3]) {
            ptr2 = 0; len2 = 0;
            throw takeFromExternrefTable0(ret[2]);
        }
        deferred3_0 = ptr2;
        deferred3_1 = len2;
        return getStringFromWasm0(ptr2, len2);
    } finally {
        wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
    }
}

/**
 * wasm wrapper over [`session::range_rects_region_by_handle`]: region-aware
 * range rects against a stored display list. `region` is
 * `"body" | "header" | "footer"`; `r_id` scopes header/footer to one HF part.
 * `Err` on an unknown/closed handle so the caller can fall back to
 * [`range_rects_region_json`].
 * @param {number} handle
 * @param {string} region
 * @param {string} r_id
 * @param {number} from
 * @param {number} to
 * @returns {string}
 */
export function range_rects_region_by_handle(handle, region, r_id, from, to) {
    let deferred4_0;
    let deferred4_1;
    try {
        const ptr0 = passStringToWasm0(region, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(r_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ret = wasm.range_rects_region_by_handle(handle, ptr0, len0, ptr1, len1, from, to);
        var ptr3 = ret[0];
        var len3 = ret[1];
        if (ret[3]) {
            ptr3 = 0; len3 = 0;
            throw takeFromExternrefTable0(ret[2]);
        }
        deferred4_0 = ptr3;
        deferred4_1 = len3;
        return getStringFromWasm0(ptr3, len3);
    } finally {
        wasm.__wbindgen_free(deferred4_0, deferred4_1, 1);
    }
}

/**
 * wasm wrapper over [`hit::range_rects_region_json`]: region-aware range rects.
 * `region` is `"body" | "header" | "footer"`; `r_id` scopes a header/footer to
 * one HF part (empty for body / match-any). The `from`/`to` refer to that
 * region's PM doc. The legacy `range_rects_json` export stays body-only.
 * @param {string} display_list
 * @param {string} region
 * @param {string} r_id
 * @param {number} from
 * @param {number} to
 * @returns {string}
 */
export function range_rects_region_json(display_list, region, r_id, from, to) {
    let deferred5_0;
    let deferred5_1;
    try {
        const ptr0 = passStringToWasm0(display_list, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(region, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ptr2 = passStringToWasm0(r_id, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len2 = WASM_VECTOR_LEN;
        const ret = wasm.range_rects_region_json(ptr0, len0, ptr1, len1, ptr2, len2, from, to);
        var ptr4 = ret[0];
        var len4 = ret[1];
        if (ret[3]) {
            ptr4 = 0; len4 = 0;
            throw takeFromExternrefTable0(ret[2]);
        }
        deferred5_0 = ptr4;
        deferred5_1 = len4;
        return getStringFromWasm0(ptr4, len4);
    } finally {
        wasm.__wbindgen_free(deferred5_0, deferred5_1, 1);
    }
}

/**
 * Register a font for measurement from raw sfnt bytes; returns the font id
 * that `measure_paragraph_json` inputs reference in their `fontChains`.
 * Malformed bytes (attacker-controlled embedded fonts) are rejected as an
 * error at this boundary, mirroring `FontStore::register`.
 * @param {Uint8Array} bytes
 * @returns {number}
 */
export function register_measure_font(bytes) {
    const ptr0 = passArray8ToWasm0(bytes, wasm.__wbindgen_malloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.register_measure_font(ptr0, len0);
    if (ret[2]) {
        throw takeFromExternrefTable0(ret[1]);
    }
    return ret[0] >>> 0;
}

/**
 * wasm wrapper over [`session::update_display_list`]: apply a page-delta
 * update to a stored display list so an incremental rebuild re-parses only
 * its changed pages. `Err` closes the handle first, so the caller's fallback
 * (a fresh [`open_display_list`]) can never race a half-updated list.
 * @param {number} handle
 * @param {string} update
 */
export function update_display_list(handle, update) {
    const ptr0 = passStringToWasm0(update, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
    const len0 = WASM_VECTOR_LEN;
    const ret = wasm.update_display_list(handle, ptr0, len0);
    if (ret[1]) {
        throw takeFromExternrefTable0(ret[0]);
    }
}
function __wbg_get_imports() {
    const import0 = {
        __proto__: null,
        __wbg___wbindgen_throw_344f42d3211c4765: function(arg0, arg1) {
            throw new Error(getStringFromWasm0(arg0, arg1));
        },
        __wbg_call_e3b662382210db98: function() { return handleError(function (arg0, arg1, arg2, arg3) {
            const ret = arg0.call(arg1, arg2, arg3);
            return ret;
        }, arguments); },
        __wbg_getRandomValues_3f44b700395062e5: function() { return handleError(function (arg0, arg1) {
            globalThis.crypto.getRandomValues(getArrayU8FromWasm0(arg0, arg1));
        }, arguments); },
        __wbg_new_from_slice_77cdfb7977362f3c: function(arg0, arg1) {
            const ret = new Uint8Array(getArrayU8FromWasm0(arg0, arg1));
            return ret;
        },
        __wbg_now_7521c72b0797ac47: function() {
            const ret = performance.now();
            return ret;
        },
        __wbindgen_cast_0000000000000001: function(arg0) {
            // Cast intrinsic for `F64 -> Externref`.
            const ret = arg0;
            return ret;
        },
        __wbindgen_cast_0000000000000002: function(arg0, arg1) {
            // Cast intrinsic for `Ref(String) -> Externref`.
            const ret = getStringFromWasm0(arg0, arg1);
            return ret;
        },
        __wbindgen_init_externref_table: function() {
            const table = wasm.__wbindgen_externrefs;
            const offset = table.grow(4);
            table.set(0, undefined);
            table.set(offset + 0, undefined);
            table.set(offset + 1, null);
            table.set(offset + 2, true);
            table.set(offset + 3, false);
        },
    };
    return {
        __proto__: null,
        "./docx_edit_bg.js": import0,
    };
}

const EditSessionFinalization = (typeof FinalizationRegistry === 'undefined')
    ? { register: () => {}, unregister: () => {} }
    : new FinalizationRegistry(ptr => wasm.__wbg_editsession_free(ptr, 1));

function addToExternrefTable0(obj) {
    const idx = wasm.__externref_table_alloc();
    wasm.__wbindgen_externrefs.set(idx, obj);
    return idx;
}

function getArrayJsValueFromWasm0(ptr, len) {
    ptr = ptr >>> 0;
    const mem = getDataViewMemory0();
    const result = [];
    for (let i = ptr; i < ptr + 4 * len; i += 4) {
        result.push(wasm.__wbindgen_externrefs.get(mem.getUint32(i, true)));
    }
    wasm.__externref_drop_slice(ptr, len);
    return result;
}

function getArrayU8FromWasm0(ptr, len) {
    ptr = ptr >>> 0;
    return getUint8ArrayMemory0().subarray(ptr / 1, ptr / 1 + len);
}

let cachedDataViewMemory0 = null;
function getDataViewMemory0() {
    if (cachedDataViewMemory0 === null || cachedDataViewMemory0.buffer.detached === true || (cachedDataViewMemory0.buffer.detached === undefined && cachedDataViewMemory0.buffer !== wasm.memory.buffer)) {
        cachedDataViewMemory0 = new DataView(wasm.memory.buffer);
    }
    return cachedDataViewMemory0;
}

function getStringFromWasm0(ptr, len) {
    return decodeText(ptr >>> 0, len);
}

let cachedUint8ArrayMemory0 = null;
function getUint8ArrayMemory0() {
    if (cachedUint8ArrayMemory0 === null || cachedUint8ArrayMemory0.byteLength === 0) {
        cachedUint8ArrayMemory0 = new Uint8Array(wasm.memory.buffer);
    }
    return cachedUint8ArrayMemory0;
}

function handleError(f, args) {
    try {
        return f.apply(this, args);
    } catch (e) {
        const idx = addToExternrefTable0(e);
        wasm.__wbindgen_exn_store(idx);
    }
}

function isLikeNone(x) {
    return x === undefined || x === null;
}

function passArray8ToWasm0(arg, malloc) {
    const ptr = malloc(arg.length * 1, 1) >>> 0;
    getUint8ArrayMemory0().set(arg, ptr / 1);
    WASM_VECTOR_LEN = arg.length;
    return ptr;
}

function passStringToWasm0(arg, malloc, realloc) {
    if (realloc === undefined) {
        const buf = cachedTextEncoder.encode(arg);
        const ptr = malloc(buf.length, 1) >>> 0;
        getUint8ArrayMemory0().subarray(ptr, ptr + buf.length).set(buf);
        WASM_VECTOR_LEN = buf.length;
        return ptr;
    }

    let len = arg.length;
    let ptr = malloc(len, 1) >>> 0;

    const mem = getUint8ArrayMemory0();

    let offset = 0;

    for (; offset < len; offset++) {
        const code = arg.charCodeAt(offset);
        if (code > 0x7F) break;
        mem[ptr + offset] = code;
    }
    if (offset !== len) {
        if (offset !== 0) {
            arg = arg.slice(offset);
        }
        ptr = realloc(ptr, len, len = offset + arg.length * 3, 1) >>> 0;
        const view = getUint8ArrayMemory0().subarray(ptr + offset, ptr + len);
        const ret = cachedTextEncoder.encodeInto(arg, view);

        offset += ret.written;
        ptr = realloc(ptr, len, offset, 1) >>> 0;
    }

    WASM_VECTOR_LEN = offset;
    return ptr;
}

function takeFromExternrefTable0(idx) {
    const value = wasm.__wbindgen_externrefs.get(idx);
    wasm.__externref_table_dealloc(idx);
    return value;
}

let cachedTextDecoder = new TextDecoder('utf-8', { ignoreBOM: true, fatal: true });
cachedTextDecoder.decode();
const MAX_SAFARI_DECODE_BYTES = 2146435072;
let numBytesDecoded = 0;
function decodeText(ptr, len) {
    numBytesDecoded += len;
    if (numBytesDecoded >= MAX_SAFARI_DECODE_BYTES) {
        cachedTextDecoder = new TextDecoder('utf-8', { ignoreBOM: true, fatal: true });
        cachedTextDecoder.decode();
        numBytesDecoded = len;
    }
    return cachedTextDecoder.decode(getUint8ArrayMemory0().subarray(ptr, ptr + len));
}

const cachedTextEncoder = new TextEncoder();

if (!('encodeInto' in cachedTextEncoder)) {
    cachedTextEncoder.encodeInto = function (arg, view) {
        const buf = cachedTextEncoder.encode(arg);
        view.set(buf);
        return {
            read: arg.length,
            written: buf.length
        };
    };
}

let WASM_VECTOR_LEN = 0;

let wasmModule, wasmInstance, wasm;
function __wbg_finalize_init(instance, module) {
    wasmInstance = instance;
    wasm = instance.exports;
    wasmModule = module;
    cachedDataViewMemory0 = null;
    cachedUint8ArrayMemory0 = null;
    wasm.__wbindgen_start();
    return wasm;
}

async function __wbg_load(module, imports) {
    if (typeof Response === 'function' && module instanceof Response) {
        if (typeof WebAssembly.instantiateStreaming === 'function') {
            try {
                return await WebAssembly.instantiateStreaming(module, imports);
            } catch (e) {
                const validResponse = module.ok && expectedResponseType(module.type);

                if (validResponse && module.headers.get('Content-Type') !== 'application/wasm') {
                    console.warn("`WebAssembly.instantiateStreaming` failed because your server does not serve Wasm with `application/wasm` MIME type. Falling back to `WebAssembly.instantiate` which is slower. Original error:\n", e);

                } else { throw e; }
            }
        }

        const bytes = await module.arrayBuffer();
        return await WebAssembly.instantiate(bytes, imports);
    } else {
        const instance = await WebAssembly.instantiate(module, imports);

        if (instance instanceof WebAssembly.Instance) {
            return { instance, module };
        } else {
            return instance;
        }
    }

    function expectedResponseType(type) {
        switch (type) {
            case 'basic': case 'cors': case 'default': return true;
        }
        return false;
    }
}

function initSync(module) {
    if (wasm !== undefined) return wasm;


    if (module !== undefined) {
        if (Object.getPrototypeOf(module) === Object.prototype) {
            ({module} = module)
        } else {
            console.warn('using deprecated parameters for `initSync()`; pass a single object instead')
        }
    }

    const imports = __wbg_get_imports();
    if (!(module instanceof WebAssembly.Module)) {
        module = new WebAssembly.Module(module);
    }
    const instance = new WebAssembly.Instance(module, imports);
    return __wbg_finalize_init(instance, module);
}

async function __wbg_init(module_or_path) {
    if (wasm !== undefined) return wasm;


    if (module_or_path !== undefined) {
        if (Object.getPrototypeOf(module_or_path) === Object.prototype) {
            ({module_or_path} = module_or_path)
        } else {
            console.warn('using deprecated parameters for the initialization function; pass a single object instead')
        }
    }

    if (module_or_path === undefined) {
        throw new Error('docx-edit wasm requires an explicit module or URL');
    }
    const imports = __wbg_get_imports();

    if (typeof module_or_path === 'string' || (typeof Request === 'function' && module_or_path instanceof Request) || (typeof URL === 'function' && module_or_path instanceof URL)) {
        module_or_path = fetch(module_or_path);
    }

    const { instance, module } = await __wbg_load(await module_or_path, imports);

    return __wbg_finalize_init(instance, module);
}

export { initSync, __wbg_init as default };
