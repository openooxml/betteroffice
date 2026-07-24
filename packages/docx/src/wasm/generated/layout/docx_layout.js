/* @ts-self-types="./docx_layout.d.ts" */

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

/**
 * @param {number} handle
 * @param {number} position
 * @param {string} direction
 * @param {number} goal_x
 * @returns {string}
 */
export function vertical_move_by_handle(handle, position, direction, goal_x) {
    let deferred3_0;
    let deferred3_1;
    try {
        const ptr0 = passStringToWasm0(direction, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.vertical_move_by_handle(handle, position, ptr0, len0, goal_x);
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
 * @param {string} display_list
 * @param {number} position
 * @param {string} direction
 * @param {number} goal_x
 * @returns {string}
 */
export function vertical_move_json(display_list, position, direction, goal_x) {
    let deferred4_0;
    let deferred4_1;
    try {
        const ptr0 = passStringToWasm0(display_list, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len0 = WASM_VECTOR_LEN;
        const ptr1 = passStringToWasm0(direction, wasm.__wbindgen_malloc, wasm.__wbindgen_realloc);
        const len1 = WASM_VECTOR_LEN;
        const ret = wasm.vertical_move_json(ptr0, len0, position, ptr1, len1, goal_x);
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
function __wbg_get_imports() {
    const import0 = {
        __proto__: null,
        __wbindgen_cast_0000000000000001: function(arg0, arg1) {
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
        "./docx_layout_bg.js": import0,
    };
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
        throw new Error('docx-layout wasm requires an explicit module or URL');
    }
    const imports = __wbg_get_imports();

    if (typeof module_or_path === 'string' || (typeof Request === 'function' && module_or_path instanceof Request) || (typeof URL === 'function' && module_or_path instanceof URL)) {
        module_or_path = fetch(module_or_path);
    }

    const { instance, module } = await __wbg_load(await module_or_path, imports);

    return __wbg_finalize_init(instance, module);
}

export { initSync, __wbg_init as default };
