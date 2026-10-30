//! wasm-bindgen session boundary over [`EditingDoc`] — the only JS-visible
//! surface of this crate, compiled behind `--features wasm`.
//!
//! Conventions every entry point below obeys, so its own docs need only state
//! what is specific to it:
//!
//! - **Values cross as JSON strings or raw bytes.** Arguments named `*_json`
//!   are JSON text; a method returning `Result<String, _>` returns JSON text.
//!   Yrs updates and binary display frames cross as byte slices / `Vec<u8>`
//!   (wasm-bindgen exposes the latter as a transferable `Uint8Array`).
//! - **Errors cross as `Err(JsValue)` carrying a display string** and nothing
//!   else — there is no error code or structured payload. Every entry point
//!   taking JSON fails that way on malformed JSON or a wrong value type, and
//!   every entry point addressing content fails that way on an unknown story,
//!   an unknown paragraph, or an out-of-range offset. A failed call leaves the
//!   document untouched: arguments are fully validated before any mutation,
//!   and each op commits in a single yrs transaction.
//! - **Content is addressed as `Loc { story, paraId, offset }`.** `offset`
//!   counts UTF-16 units within ONE paragraph and lives in `[0, para_len]`,
//!   where `para_len` excludes the paragraph's own pilcrow; `offset ==
//!   para_len` addresses the paragraph mark. A paragraph id is resolved
//!   story-scoped, so an id living in another story reads as "not found".
//!   Story-global indices exist only transiently inside a call and are never
//!   returned.
//! - **A range is two Locs in one story** and is half-open, `[start, end)`.
//!   Because pilcrows occupy a unit, a range whose ends sit in different
//!   paragraphs spans the boundary marks.
//! - **Suggesting mode is the optional `(author_name, author_date)` pair.**
//!   Pass both to record the edit as a tracked change stamped with that author
//!   and ISO date, or neither for a plain local edit; passing one alone is an
//!   error. Ops without the pair are always plain and local.
//! - **Receipts are JSON objects.** An op that can stamp a tracked change
//!   reports `{"revisionId": string|null}` — null outside suggesting mode.
//!   Ops with no receipt content return `()`.
//!
//! - **Each session owns its measurement fonts.** Every entry point that
//!   reaches font registration, measurement, layout, display state or glyph
//!   outlines starts with `let _fonts = self.core.fonts.enter();`
//!   ([`docx_layout::MeasureFonts`]), so font ids are the session's own. One
//!   that forgets panics on its first font lookup.
//!
//! Story lengths, selection indices and every other unit count in this module
//! are UTF-16 units in which each embed, pilcrows included, counts as one.

use std::sync::Arc;

use js_sys::{Function, Uint8Array};
#[cfg(test)]
use serde_json::json;
use wasm_bindgen::prelude::*;
use yrs::Subscription;

use crate::session::{EditSessionCore, EngineError, EngineErrorKind, Host};

fn js_err(error: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&error.to_string())
}

impl From<EngineError> for JsValue {
    fn from(error: EngineError) -> Self {
        match error.kind {
            EngineErrorKind::PeerMetadata { code } => {
                let value = js_sys::Error::new(&error.message);
                value.set_name("PeerMetadataError");
                let _ = js_sys::Reflect::set(
                    &value,
                    &JsValue::from_str("code"),
                    &JsValue::from_str(code),
                );
                value.into()
            }
            EngineErrorKind::Message => JsValue::from_str(&error.message),
        }
    }
}

fn wasm_host() -> Host {
    #[wasm_bindgen]
    extern "C" {
        #[wasm_bindgen(js_namespace = performance, js_name = now)]
        fn performance_now() -> f64;
    }

    Host::new(
        Arc::new(|| js_sys::Date::now() as u64),
        Arc::new(js_entropy),
        Arc::new(performance_now),
    )
}

/// Host randomness for version nonces; the wasm target has no ambient entropy source.
fn js_entropy() -> u64 {
    #[cfg(target_arch = "wasm32")]
    {
        ((js_sys::Math::random() * 9_007_199_254_740_992.0) as u64)
            ^ (js_sys::Date::now() as u64).rotate_left(21)
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        0
    }
}
#[wasm_bindgen]
pub struct RetainedLayoutMeta {
    inner: crate::engine::RetainedLayoutMeta,
}

#[wasm_bindgen]
impl RetainedLayoutMeta {
    #[wasm_bindgen(getter)]
    pub fn page_count(&self) -> u32 {
        self.inner.page_count as u32
    }

    #[wasm_bindgen(getter)]
    pub fn partial(&self) -> bool {
        self.inner.partial
    }

    #[wasm_bindgen(getter)]
    pub fn provisional(&self) -> bool {
        self.inner.provisional
    }

    #[wasm_bindgen(getter)]
    pub fn notes_converged(&self) -> bool {
        self.inner.notes_converged
    }

    pub fn page_sizes(&self) -> Vec<f64> {
        self.inner.page_sizes.clone()
    }

    pub fn layout_shell_json(&self) -> String {
        self.inner.layout_shell_json.clone()
    }
}

/// What runs inherit under one package's styles; see [`EditSession::run_inheritance`].
#[wasm_bindgen]
pub struct RunInheritance {
    inner: crate::seed::RunInheritance,
}

#[wasm_bindgen]
impl RunInheritance {
    /// The attributes, as a JSON object, that seeding merges into the marks of a run of
    /// `run_style` in a paragraph of `paragraph_style` before the run's direct formatting.
    pub fn attributes(
        &mut self,
        paragraph_style: Option<String>,
        run_style: Option<String>,
        table_run_json: Option<String>,
    ) -> Result<String, JsValue> {
        let attrs = if let Some(table_run_json) = table_run_json {
            self.inner.table_run_attrs(
                paragraph_style.as_deref(),
                run_style.as_deref(),
                serde_json::from_str(&table_run_json).map_err(js_err)?,
            )
        } else {
            self.inner
                .run_attrs(paragraph_style.as_deref(), run_style.as_deref())
        };
        serde_json::to_string(&attrs).map_err(js_err)
    }

    pub fn paragraph_formatting(
        &mut self,
        paragraph_style: Option<String>,
        table_formatting_json: &str,
    ) -> Result<String, JsValue> {
        let table_formatting = serde_json::from_str(table_formatting_json).map_err(js_err)?;
        let formatting = self
            .inner
            .paragraph_formatting(paragraph_style.as_deref(), table_formatting);
        serde_json::to_string(&formatting).map_err(js_err)
    }

    pub fn table_attributes(&self, table_json: &str) -> Result<String, JsValue> {
        let table = serde_json::from_str(table_json).map_err(js_err)?;
        serde_json::to_string(&self.inner.table_attrs(&table)).map_err(js_err)
    }
}

/// One yrs replica of the DOCX editing model, held for a JS host.
///
/// Owns the [`EditingDoc`] plus the (single) JS update observer. The JS facade
/// multiplexes its own listener set over that one callback.
#[wasm_bindgen]
pub struct EditSession {
    core: EditSessionCore,
    update_observer: Option<Subscription>,
}

#[wasm_bindgen]
impl EditSession {
    /// Creates a replica. The host allocates `client_id` (a random 32-bit id
    /// is fine) and must keep it unique across the replicas that will merge.
    /// Errors unless it is a non-negative safe integer.
    #[wasm_bindgen(constructor)]
    pub fn new(client_id: f64) -> Result<EditSession, JsValue> {
        const MAX_SAFE_INTEGER: f64 = 9_007_199_254_740_991.0;
        if !(client_id.is_finite()
            && client_id >= 0.0
            && client_id.fract() == 0.0
            && client_id <= MAX_SAFE_INTEGER)
        {
            return Err(js_err("client_id must be a non-negative safe integer"));
        }
        Ok(Self {
            core: EditSessionCore::new(client_id as u64, wasm_host()),
            update_observer: None,
        })
    }

    /// This replica's client id, as passed to the constructor.
    pub fn client_id(&self) -> f64 {
        self.core.client_id()
    }

    /// Registers raw sfnt bytes in this session's resident measurement store
    /// and returns the font id that measurement and display inputs reference.
    /// Errors on bytes the font parser rejects.
    pub fn register_measure_font(&self, bytes: &[u8]) -> Result<u32, JsValue> {
        self.core
            .register_measure_font(bytes)
            .map_err(JsValue::from)
    }

    pub fn decode_font(&self, bytes: &[u8]) -> Result<Vec<u8>, JsValue> {
        self.core.decode_font(bytes).map_err(JsValue::from)
    }

    /// Registers a measurement view of `base` carrying the vertical metrics
    /// and advance pitch Word measures `requested_family` with — for a face
    /// this host had to substitute. Returns `base` for a family whose metrics
    /// are unknown.
    pub fn register_substitute_measure_font(
        &self,
        base: u32,
        requested_family: &str,
    ) -> Result<u32, JsValue> {
        self.core
            .register_substitute_measure_font(base, requested_family)
            .map_err(JsValue::from)
    }

    /// Drops every registered measurement font (ids restart at zero) and
    /// invalidates the retained paragraph measurement templates, so the next
    /// edit must pass back through the full layout path.
    pub fn clear_measure_fonts(&self) {
        self.core.clear_measure_fonts()
    }

    /// Measurement input JSON in, `ParagraphExtent` JSON out. Also records the
    /// paragraph's immutable width/font envelope under its stable block id, so
    /// a later resident edit re-measures only the changed block. Errors with
    /// the engine's message for input it cannot measure.
    pub fn measure_paragraph_json(&self, input: &str) -> Result<String, JsValue> {
        self.core
            .measure_paragraph_json(input)
            .map_err(JsValue::from)
    }

    /// `{ measured, options }` JSON in, `Layout` JSON out. Both the measured
    /// input and the resulting layout are retained for the resident edit path.
    /// Errors on unparseable input or on layout failure.
    pub fn layout_document_json(&self, input: &str) -> Result<String, JsValue> {
        self.core.layout_document_json(input).map_err(JsValue::from)
    }

    /// Region-layout input JSON in, the font families and sizes that input
    /// needs as JSON out, so the host can register fonts before laying out.
    pub fn layout_font_requirements_json(&self, input: &str) -> Result<String, JsValue> {
        self.core
            .layout_font_requirements_json(input)
            .map_err(JsValue::from)
    }

    /// Region-layout input JSON in, a paginated envelope with section and page
    /// regions already composed as JSON out — ready for the display builder
    /// with no host-side layout mutation. Retains the pass for resident edits.
    pub fn layout_document_with_regions_json(&self, input: &str) -> Result<String, JsValue> {
        self.core
            .layout_document_with_regions_json(input)
            .map_err(JsValue::from)
    }

    /// Same full region pass as [`Self::layout_document_with_regions_json`],
    /// but the reply carries only `{ layout, headersFooters?, notesConverged }`
    /// — the measured arena stays retained wasm-side and is fetched on demand
    /// via [`Self::retained_kernel_inputs_json`].
    pub fn layout_document_with_regions_retained_json(
        &self,
        input: &str,
    ) -> Result<String, JsValue> {
        self.core
            .layout_document_with_regions_retained_json(input)
            .map_err(JsValue::from)
    }

    pub fn layout_document_with_regions_retained_meta(
        &self,
        input: &str,
    ) -> Result<RetainedLayoutMeta, JsValue> {
        self.core
            .layout_document_with_regions_retained_meta(input)
            .map(|inner| RetainedLayoutMeta { inner })
            .map_err(JsValue::from)
    }

    pub fn retained_layout_json(&self) -> Result<String, JsValue> {
        self.core.retained_layout_json().map_err(JsValue::from)
    }

    /// [`Self::layout_document_with_regions_retained_json`] without the reply,
    /// for a caller that reads only the retained state.
    pub fn layout_document_with_regions_retained(&self, input: &str) -> Result<(), JsValue> {
        self.core
            .layout_document_with_regions_retained(input)
            .map_err(JsValue::from)
    }

    /// [`Self::layout_document_with_regions_retained`] after one edit: the
    /// resident path when that edit was the [`Self::edit_text_resident`] taken
    /// last; see `EngineSession::layout_document_with_regions_after_edit`.
    pub fn layout_document_with_regions_after_edit(&self, input: &str) -> Result<(), JsValue> {
        self.core
            .layout_document_with_regions_after_edit(input)
            .map_err(JsValue::from)
    }

    /// The retained region layout's `headersFooters` JSON, when it has any.
    pub fn retained_headers_footers_json(&self) -> Result<Option<String>, JsValue> {
        self.core
            .retained_headers_footers_json()
            .map_err(JsValue::from)
    }

    /// The retained region layout of the first `pages` pages only; see
    /// `EngineSession::layout_document_with_regions_prefix_retained_json`.
    pub fn layout_document_with_regions_prefix_retained_json(
        &self,
        input: &str,
        pages: u32,
    ) -> Result<String, JsValue> {
        self.core
            .layout_document_with_regions_prefix_retained_json(input, pages)
            .map_err(JsValue::from)
    }

    /// The retained region layout of exactly the first `pages` pages when no
    /// later content can change them, and of the whole document otherwise; see
    /// `EngineSession::layout_document_with_regions_exact_prefix_retained`.
    /// Returns `{ provisional, pages, documentWidth? }` as JSON.
    pub fn layout_document_with_regions_exact_prefix_retained(
        &self,
        input: &str,
        pages: u32,
    ) -> Result<String, JsValue> {
        self.core
            .layout_document_with_regions_exact_prefix_retained(input, pages)
            .map_err(JsValue::from)
    }

    /// Begins `layout_document_with_regions_retained_json` as a pass measured a
    /// step at a time; see `EngineSession::begin_region_layout`. Returns the
    /// progress JSON, with `layoutJson` once the pass is complete.
    pub fn begin_region_layout(&self, input: &str) -> Result<String, JsValue> {
        self.core.begin_region_layout(input).map_err(JsValue::from)
    }

    /// Measures up to `blocks` more body blocks of the begun pass; see
    /// `EngineSession::resume_region_layout`.
    pub fn resume_region_layout(&self, blocks: u32) -> Result<String, JsValue> {
        self.core
            .resume_region_layout(blocks)
            .map_err(JsValue::from)
    }

    /// Retained `{ measured, options }` for the main-thread display-list
    /// fallback after a retained-only region layout.
    pub fn retained_kernel_inputs_json(&self) -> Result<String, JsValue> {
        self.core
            .retained_kernel_inputs_json()
            .map_err(JsValue::from)
    }

    /// Makes the next display frame a full one whatever epoch its caller
    /// passes, for a host that switches to this engine from another.
    pub fn reset_frame_base(&self) {
        self.core.reset_frame_base()
    }

    /// `{ measured, options, layout }` JSON in, `DisplayList` JSON out, built
    /// against the same resident font store this session measures with.
    pub fn build_display_list_json(&self, input: &str) -> Result<String, JsValue> {
        self.core
            .build_display_list_json(input)
            .map_err(JsValue::from)
    }

    /// Display-only input JSON in, one binary `FrameDelta` v1 out (exposed as
    /// a transferable `Uint8Array`). `expected_frame_epoch` is the epoch of the
    /// frame the caller currently holds; pass `0` for the first frame. A
    /// mismatch makes the engine emit a full frame instead of a delta, and the
    /// returned frame's epoch is always greater than `expected_frame_epoch`. Errors
    /// unless the epoch is a non-negative safe integer, and on build failure.
    pub fn build_display_list_frame(
        &self,
        input: &str,
        expected_frame_epoch: f64,
    ) -> Result<Vec<u8>, JsValue> {
        self.core
            .build_display_list_frame(input, expected_frame_epoch)
            .map_err(JsValue::from)
    }

    /// Limit full display builds to pages `start..end` plus the pages already
    /// built; the others stay unbuilt placeholders carrying their geometry
    /// until [`Self::build_display_pages_frame`] builds them.
    pub fn set_display_window(&self, start: u32, end: u32) {
        self.core.set_display_window(start, end)
    }

    /// Keep every previously built page while windowed builds are on.
    pub fn set_display_retain_built_pages(&self, retain: bool) {
        self.core.set_display_retain_built_pages(retain)
    }

    /// Let an eligible resident text edit re-lower only its paragraph. Off by default.
    pub fn set_local_lowering(&self, enabled: bool) {
        self.core.set_local_lowering(enabled)
    }

    /// Apply admitted paragraph-local text batches directly. Off by default.
    pub fn set_direct_batches(&self, enabled: bool) {
        self.core.set_direct_batches(enabled)
    }

    /// The number of text batches applied directly to the live document.
    pub fn direct_batches_applied(&self) -> f64 {
        self.core.direct_batches_applied()
    }

    /// Limit incremental rebuilds to the display window and caret pages. Off by default.
    pub fn set_windowed_incremental_builds(&self, enabled: bool) {
        self.core.set_windowed_incremental_builds(enabled)
    }

    /// Build the listed pages that are still unbuilt and return a FrameDelta
    /// v1 carrying them; `expected_frame_epoch` works as for
    /// [`Self::build_display_list_frame`].
    pub fn build_display_pages_frame(
        &self,
        pages: Vec<u32>,
        expected_frame_epoch: f64,
    ) -> Result<Vec<u8>, JsValue> {
        self.core
            .build_display_pages_frame(pages, expected_frame_epoch)
            .map_err(JsValue::from)
    }

    /// Release display pages; an empty result means the request was superseded.
    pub fn release_display_pages_frame(
        &self,
        pages: Vec<u32>,
        expected_frame_epoch: f64,
    ) -> Result<Vec<u8>, JsValue> {
        self.core
            .release_display_pages_frame(pages, expected_frame_epoch)
            .map_err(JsValue::from)
    }

    /// `{"frameEpoch", "caretRect": {…}|null}` for the session's own collapsed
    /// body selection. `caretRect` is null whenever there is no selection, the
    /// selection is not a collapsed body caret, or the retained layout has no
    /// geometry for it. `frameEpoch` identifies the frame the rect belongs to.
    pub fn resident_caret_snapshot_json(&self) -> Result<String, JsValue> {
        self.core
            .resident_caret_snapshot_json()
            .map_err(JsValue::from)
    }

    /// Applies one ordinary insertion at this session's collapsed selection
    /// and returns the resulting binary `FrameDelta`. The inserted text
    /// inherits the formatting at the caret; selection, measurement inputs,
    /// pagination checkpoints and display state all stay resident, so nothing
    /// but the frame crosses the boundary.
    ///
    /// Errors when `text` is empty or contains `\r`/`\n`, when
    /// `expected_frame_epoch` is not a non-negative safe integer, when there is
    /// no resident selection or it no longer resolves, when the selection is
    /// not collapsed, or when the resident layout state cannot absorb an edit
    /// in that paragraph. The caller must then run the full layout path.
    pub fn apply_input(&self, text: &str, expected_frame_epoch: f64) -> Result<Vec<u8>, JsValue> {
        self.core
            .apply_input(text, expected_frame_epoch)
            .map_err(JsValue::from)
    }

    /// Instrumented twin of [`EditSession::apply_input`]: identical arguments,
    /// result and error contract, but it also records stage timings for
    /// [`EditSession::apply_input_profile_json`]. Separate so the ordinary
    /// input path pays no timer calls.
    pub fn apply_input_profiled(
        &self,
        text: &str,
        expected_frame_epoch: f64,
    ) -> Result<Vec<u8>, JsValue> {
        self.core
            .apply_input_profiled(text, expected_frame_epoch)
            .map_err(JsValue::from)
    }

    /// Deletes up to `count` characters at this session's collapsed selection,
    /// lays out once, and returns the resulting binary `FrameDelta`.
    /// `direction` is `"backward"` or `"forward"`; a surrogate pair is removed
    /// whole. At a paragraph boundary a deletion merges with the neighbouring
    /// paragraph instead. Deleting stops early at the document start or end,
    /// before a second paragraph merge, or at a paragraph the resident state
    /// cannot absorb; [`EditSession::resident_deleted_units`] reports how many
    /// were removed.
    ///
    /// Errors on an unknown `direction`, a zero `count`, when
    /// `expected_frame_epoch` is not a non-negative safe integer, under the
    /// same selection and readiness conditions as [`EditSession::apply_input`],
    /// and when there is no character to delete in that direction.
    pub fn apply_delete(
        &self,
        direction: &str,
        expected_frame_epoch: f64,
        count: u32,
    ) -> Result<Vec<u8>, JsValue> {
        self.core
            .apply_delete(direction, expected_frame_epoch, count)
            .map_err(JsValue::from)
    }

    /// Characters the last [`EditSession::apply_delete`] removed.
    pub fn resident_deleted_units(&self) -> u32 {
        self.core.resident_deleted_units()
    }

    /// Instrumented twin of [`EditSession::apply_delete`]: identical arguments,
    /// result and error contract, but it also records stage timings for
    /// [`EditSession::apply_input_profile_json`].
    pub fn apply_delete_profiled(
        &self,
        direction: &str,
        expected_frame_epoch: f64,
        count: u32,
    ) -> Result<Vec<u8>, JsValue> {
        self.core
            .apply_delete_profiled(direction, expected_frame_epoch, count)
            .map_err(JsValue::from)
    }

    /// Stage timings of the last profiled apply, in milliseconds:
    /// `{"selectionMs","editMs","lowerMs","measureMs","paginateMs",
    /// "displayInputMs","displayBuildMs","displayFinalizeMs","displayMs",
    /// "encodeMs"}`. `{}` before the first profiled call.
    pub fn apply_input_profile_json(&self) -> String {
        self.core.apply_input_profile_json()
    }

    /// Region-aware hit test against the resident display list, so no
    /// display-list JSON crosses the boundary. `x`/`y` are page-local px.
    /// Returns
    /// `{"region":"body"|"header"|"footer"|"footnote"|"endnote","rId"?,`
    /// `"noteId"?,"pos":n|null,"target":"text"|"image"|"none"}`, or `"null"`
    /// for a page index outside the frame. A header/footer `pos` refers to
    /// that part's own document and a note `pos` to the `fn:{id}` / `en:{id}`
    /// story, never the body; `target` is what sits under the point, which is
    /// what a pointer cursor needs.
    pub fn display_hit_test_regions_json(
        &self,
        page_index: u32,
        x: f64,
        y: f64,
    ) -> Result<String, JsValue> {
        self.core
            .display_hit_test_regions_json(page_index, x, y)
            .map_err(JsValue::from)
    }

    /// Nearest caret position on the adjacent visual line. `direction` is
    /// `"up"` or `"down"`; `goal_x` is the page-local x the caret is trying to
    /// hold across successive moves (a non-finite value falls back to the
    /// caret's own x). Returns `{"position","goalX"}`, or `"null"` when there
    /// is no line in that direction. Errors on an unknown `direction` and when
    /// no display list is resident.
    pub fn display_vertical_move_json(
        &self,
        position: f64,
        direction: &str,
        goal_x: f64,
    ) -> Result<String, JsValue> {
        self.core
            .display_vertical_move_json(position, direction, goal_x)
            .map_err(JsValue::from)
    }

    /// Highlight rectangles for a body document range against the resident
    /// display list: a JSON array of `{pageIndex,x,y,width,height}` in
    /// page-local px, one entry per page the range touches. Body positions
    /// only. Errors when no display list is resident.
    pub fn display_range_rects_json(&self, from: f64, to: f64) -> Result<String, JsValue> {
        self.core
            .display_range_rects_json(from, to)
            .map_err(JsValue::from)
    }

    /// @internal
    pub fn display_range_rects_on_pages_json(
        &self,
        from: f64,
        to: f64,
        first_page: f64,
        last_page: f64,
    ) -> Result<String, JsValue> {
        self.core
            .display_range_rects_on_pages_json(from, to, first_page, last_page)
            .map_err(JsValue::from)
    }

    /// Same rectangles as [`EditSession::display_range_rects_json`], scoped to
    /// a region. `region` is `"body"`, `"header"`, `"footer"`, `"footnote"` or
    /// `"endnote"`; `part_id` names one header/footer part (an empty string
    /// matches any) or the note whose story the positions belong to.
    /// `from`/`to` are positions in THAT region's document, and a
    /// header/footer part paints on every page carrying it, so the result
    /// holds one rect set per such page, each tagged with its `pageIndex`.
    /// Errors on an unknown `region` and when no display list is resident.
    pub fn display_range_rects_region_json(
        &self,
        region: &str,
        part_id: &str,
        from: f64,
        to: f64,
    ) -> Result<String, JsValue> {
        self.core
            .display_range_rects_region_json(region, part_id, from, to)
            .map_err(JsValue::from)
    }

    /// [`EditSession::display_range_rects_region_json`] for many ranges, given as flat
    /// `[from, to, from, to, ...]` pairs. Per range, in order: its rect count, then
    /// `pageIndex, x, y, width, height` for each rect.
    pub fn display_range_rects_batch(
        &self,
        region: &str,
        part_id: &str,
        ranges: Vec<f64>,
    ) -> Result<Vec<f64>, JsValue> {
        self.core
            .display_range_rects_batch(region, part_id, ranges)
            .map_err(JsValue::from)
    }

    /// One glyph outline from this session's resident font store:
    /// `{"upem":n,"cmds":[{"t":"M"|"L"|"Q"|"C"|"Z", …}]}` — commands in font
    /// units, y-up. `font_id` comes from
    /// [`EditSession::register_measure_font`] and `glyph_id` from shaping.
    /// Errors when the glyph cannot be extracted.
    pub fn outline_glyph_json(&self, font_id: u32, glyph_id: u32) -> Result<String, JsValue> {
        self.core
            .outline_glyph_json(font_id, glyph_id)
            .map_err(JsValue::from)
    }

    /// Typing echo model of one font at one shaping size from this session's
    /// resident font store: `{"horizontal","glyphIds","advances","kerning",
    /// "unsafePairs"}`, indexed by the characters of `alphabet`. `kern` is the
    /// run's kerning gate. Errors for an unknown font or a non-positive size.
    pub fn echo_glyph_set_json(
        &self,
        font_id: u32,
        size: f32,
        kern: bool,
        alphabet: &str,
    ) -> Result<String, JsValue> {
        self.core
            .echo_glyph_set_json(font_id, size, kern, alphabet)
            .map_err(JsValue::from)
    }

    /// Hydrates from a yrs v1 update. The first load after an unseeded open retains prior
    /// comment writes and marks loaded fields differing from seed placeholders as authored.
    pub fn load(&self, update: &[u8]) -> Result<(), JsValue> {
        self.core.load(update).map_err(JsValue::from)
    }

    pub fn encode_peer_metadata(&self) -> Result<Vec<u8>, JsValue> {
        self.core.encode_peer_metadata().map_err(JsValue::from)
    }

    pub fn bootstrap_peer(
        &self,
        state: &[u8],
        metadata: &[u8],
        source: Option<Vec<u8>>,
    ) -> Result<(), JsValue> {
        self.core
            .bootstrap_peer(state, metadata, source)
            .map_err(JsValue::from)
    }

    /// [`EditSession::open_docx`] with seeding always on.
    #[wasm_bindgen(js_name = seed_from_docx)]
    pub fn seed_from_docx_owned(
        &self,
        bytes: Vec<u8>,
        generation: Option<String>,
    ) -> Result<String, JsValue> {
        self.core
            .seed_from_docx_owned(bytes, generation)
            .map_err(JsValue::from)
    }

    /// Starts a new opening of the document; see [`EditingDoc::begin_opening`].
    pub fn begin_opening(&self, generation: Option<String>) {
        self.core.begin_opening(generation)
    }

    /// Declares header/footer alias groups before host-side seeding.
    pub fn set_header_footer_aliases(&self, json: &str) -> Result<(), JsValue> {
        self.core
            .set_header_footer_aliases(json)
            .map_err(JsValue::from)
    }

    /// Returns active alias-to-canonical relationship ids as a JSON object.
    pub fn header_footer_aliases_json(&self) -> String {
        self.core.header_footer_aliases_json()
    }

    /// Unions seeded opaque sequence names into document state.
    pub fn seed_opaque_sequences(&self, names_json: &str) -> Result<(), JsValue> {
        self.core
            .seed_opaque_sequences(names_json)
            .map_err(JsValue::from)
    }

    /// Parses a DOCX package, optionally seeds its editable stories into this
    /// replica, and retains the source bytes for
    /// [`EditSession::materialize_docx`] and paragraph identity reads.
    /// Seeding starts a new opening, with `generation` or a fresh one, so its
    /// session anchors are its own; see [`EditingDoc::begin_opening`].
    ///
    /// Returns `{"envelope","referencedFonts":[string, …],"unusedScriptFonts":[string, …]}`:
    /// `unusedScriptFonts` are the referenced fonts a seeded package names only
    /// for East Asian or complex-script text it does not contain. The envelope is
    /// the parsed package with the parts the host does not need stripped —
    /// body content, header/footer and note content, numbering, media and
    /// charts are emptied, section entries keep only their properties — so
    /// styles, theme, settings, fonts and relationships still cross while the
    /// bulk of the document stays in Rust. Errors on bytes that are not a
    /// readable DOCX.
    /// `digest`, when given, must be the SHA-256 of `bytes` in lowercase hex,
    /// as a host that hashed them off this thread already knows it.
    #[wasm_bindgen(js_name = open_docx)]
    pub fn open_docx_owned(
        &self,
        bytes: Vec<u8>,
        seed_stories: bool,
        generation: Option<String>,
        digest: Option<String>,
    ) -> Result<String, JsValue> {
        self.core
            .open_docx_owned(bytes, seed_stories, generation, digest)
            .map_err(JsValue::from)
    }

    /// Opens `bytes` for display only, seeded from the body's first `blocks`
    /// blocks (see `seed::seed_docx_preview`): the reply is the host metadata
    /// of that parse, as [`EditSession::open_docx`] replies, with the
    /// `unusedScriptFonts` of its cut. The session keeps no source package, so
    /// it cannot save.
    /// Opens nothing and replies with nothing for a document the preview
    /// refuses, which opens with [`EditSession::open_docx`] instead.
    #[wasm_bindgen(js_name = open_docx_preview)]
    pub fn open_docx_preview_owned(
        &self,
        bytes: Vec<u8>,
        blocks: u32,
    ) -> Result<Option<String>, JsValue> {
        self.core
            .open_docx_preview_owned(bytes, blocks)
            .map_err(JsValue::from)
    }

    #[wasm_bindgen(js_name = open_docx_preview_with_budget)]
    pub fn open_docx_preview_with_budget_owned(
        &self,
        bytes: Vec<u8>,
        blocks: u32,
        paragraph_budget: Option<u32>,
    ) -> Result<Option<String>, JsValue> {
        self.core
            .open_docx_preview_with_budget_owned(bytes, blocks, paragraph_budget)
            .map_err(JsValue::from)
    }

    /// Whether [`EditSession::open_docx`] seeds images as `media:{n}` tokens,
    /// which only a replica opened from the same package resolves, instead of
    /// `data:` URLs. Off by default.
    pub fn set_media_tokens(&self, enabled: bool) {
        self.core.set_media_tokens(enabled)
    }

    /// The digests mapping the `data:` URLs this replica seeded to
    /// `media:{n}` tokens, for [`EditSession::load_media_sources`]; empty
    /// when it seeded none.
    pub fn media_sources_json(&self) -> String {
        self.core.media_sources_json()
    }

    /// Lays this replica's `data:` image sources out as the `media:{n}` tokens
    /// another replica seeded them from.
    pub fn load_media_sources(&self, json: &str) -> Result<(), JsValue> {
        self.core.load_media_sources(json).map_err(JsValue::from)
    }

    /// The separator notes for another replica; empty when the package has none.
    pub fn note_separators_state(&self) -> Result<Vec<u8>, JsValue> {
        self.core.note_separators_state().map_err(JsValue::from)
    }

    /// Loads separator notes from a yrs v1 update; empty clears them.
    pub fn load_note_separators(&self, state: &[u8]) -> Result<(), JsValue> {
        self.core.load_note_separators(state).map_err(JsValue::from)
    }

    /// Marks whether the session's document is part of a package, as a
    /// replica of a preview is: its layouts render NUMPAGES empty.
    pub fn set_partial_document(&self, partial: bool) {
        self.core.set_partial_document(partial)
    }

    /// Re-parses the DOCX bytes retained by the last
    /// [`EditSession::open_docx`] and returns the COMPLETE package envelope as
    /// JSON, or `None` when no DOCX has been opened. Unlike `open_docx` this
    /// keeps every part; it reflects the source file, not later edits.
    pub fn materialize_docx(&self) -> Result<Option<String>, JsValue> {
        self.core.materialize_docx().map_err(JsValue::from)
    }

    /// The bytes a `media:{n}` image source displays, as the part's
    /// [`EditSession::media_type`]; `undefined` for any other source or a
    /// part that cannot be read.
    pub fn media_bytes(&self, token: &str) -> Option<Vec<u8>> {
        self.core.media_bytes(token)
    }

    /// The media type of [`EditSession::media_bytes`].
    pub fn media_type(&self, token: &str) -> Option<String> {
        self.core.media_type(token)
    }

    /// The `data:` URL a `media:{n}` image source stands for.
    pub fn media_data_url(&self, token: &str) -> Option<String> {
        self.core.media_data_url(token)
    }

    /// Seeds stories from JSON:
    /// `[{"storyId","paragraphs":[{"text","pStyle"?,"alignment"?}, …]}, …]`.
    /// `pStyle` defaults to `"Normal"` and `alignment` to `"left"`; paragraph
    /// text must not contain paragraph breaks. Receipt:
    /// `{storyId: [paraId, …]}` with each story's paragraphs in document
    /// order. Errors when a story entry has no `storyId`, no `paragraphs`
    /// array, or an empty one, and when a story id already exists. Seeding a
    /// document that has no opening yet starts one; see
    /// [`EditingDoc::begin_opening`].
    pub fn load_json(&self, stories_json: &str) -> Result<String, JsValue> {
        self.core.load_json(stories_json).map_err(JsValue::from)
    }

    /// The full document state as one yrs v1 update. Hand it to
    /// [`EditSession::load`] on a fresh replica to reproduce this document.
    pub fn encode_state(&self) -> Vec<u8> {
        self.core.encode_state()
    }

    /// This replica's yrs v1 state vector — what a peer sends so
    /// [`EditSession::encode_diff`] can compute the update it is missing.
    pub fn encode_state_vector(&self) -> Vec<u8> {
        self.core.encode_state_vector()
    }

    /// Whether a remote update is waiting on one that has not arrived yet.
    pub fn has_pending_update(&self) -> bool {
        self.core.has_pending_update()
    }

    /// The yrs v1 update carrying everything this replica has that the peer
    /// described by `remote_state_vector` does not. Errors on a malformed
    /// state vector.
    pub fn encode_diff(&self, remote_state_vector: &[u8]) -> Result<Vec<u8>, JsValue> {
        self.core
            .encode_diff(remote_state_vector)
            .map_err(JsValue::from)
    }

    /// [`EditSession::encode_diff`] as ordered updates for a sync that must
    /// travel in bounded frames: the whole diff when it fits
    /// `max_chunk_bytes`, otherwise causally ordered struct updates followed
    /// by the delete set alone. Each fits the limit except one that carries a
    /// single larger struct. Errors on a malformed state vector or a zero
    /// limit.
    pub fn encode_diff_chunks(
        &self,
        remote_state_vector: &[u8],
        max_chunk_bytes: u32,
    ) -> Result<js_sys::Array, JsValue> {
        let chunks = self
            .core
            .encode_diff_chunks(remote_state_vector, max_chunk_bytes)?;
        Ok(chunks
            .iter()
            .map(|chunk| Uint8Array::from(chunk.as_slice()))
            .collect())
    }

    /// Applies a remote or incremental yrs v1 update. It commits without a
    /// local origin, so this replica's undo manager never claims it. Errors on
    /// a malformed update.
    pub fn apply_update(&self, update: &[u8]) -> Result<(), JsValue> {
        self.core.apply_update(update).map_err(JsValue::from)
    }

    /// Applies the ordered updates of one sync in a single remote transaction,
    /// so the result matches [`EditSession::apply_update`] on the whole diff.
    /// Errors on a non-`Uint8Array` entry or a malformed or oversized update,
    /// before applying any.
    pub fn apply_update_chunks(&self, updates: js_sys::Array) -> Result<(), JsValue> {
        let chunks = updates
            .iter()
            .map(|update| {
                update
                    .dyn_into::<Uint8Array>()
                    .map(|bytes| bytes.to_vec())
                    .map_err(|_| js_err("sync updates must be Uint8Arrays"))
            })
            .collect::<Result<Vec<_>, _>>()?;
        self.core
            .apply_update_chunks(&chunks)
            .map_err(JsValue::from)
    }

    /// [`EditSession::apply_update`] plus an inference about where the peer
    /// that authored it is typing, for caret presence. Returns
    /// `{"clientId","story","paraId","endOffset"}` for the end of that peer's
    /// last text insertion, or `"null"` when the update carries work from more
    /// than one client, ends in a non-text insertion, or inserts nothing.
    /// Errors on a malformed update.
    pub fn apply_update_with_inference(&self, update: &[u8]) -> Result<String, JsValue> {
        self.core
            .apply_update_with_inference(update)
            .map_err(JsValue::from)
    }

    /// Applies an update produced by this document's dedicated local worker.
    /// The local origin lets the main replica's UndoManager retain ownership of
    /// the edit; remote/collaboration updates must use `apply_update` instead.
    pub fn apply_local_update(&self, update: &[u8]) -> Result<(), JsValue> {
        self.core.apply_local_update(update).map_err(JsValue::from)
    }

    /// Applies an update another replica of this document committed for a host batch.
    /// It commits outside undo history and notifies as a local change.
    pub fn apply_host_update(&self, update: &[u8]) -> Result<(), JsValue> {
        self.core.apply_host_update(update).map_err(JsValue::from)
    }

    /// Subscribes `callback(update: Uint8Array, isRemote: 0|1)` to every
    /// committed transaction. `update` is v1-encoded — feed it straight to
    /// [`EditSession::apply_update`] on a peer — and is copied out of wasm
    /// memory, so JS may hold it across later edits. The second argument is
    /// `1` when the transaction had no local origin; a remote update's own
    /// formatting cleanup follows it as a `0` update to forward. One observer
    /// per session; a second call replaces the first, and a throwing callback
    /// is ignored.
    pub fn set_update_observer(&mut self, callback: &Function) -> Result<(), JsValue> {
        let callback = callback.clone();
        let subscription = self
            .core
            .engine
            .doc()
            .observe_update_v1(move |update, remote| {
                // Uint8Array::from copies out of wasm memory — the JS side
                // owns the bytes and may hold them across further edits.
                let bytes = Uint8Array::from(update);
                let origin = if remote { 1.0 } else { 0.0 };
                let _ = callback.call2(&JsValue::NULL, &bytes.into(), &JsValue::from_f64(origin));
            })
            .map_err(js_err)?;
        self.update_observer = Some(subscription);
        Ok(())
    }

    /// Drops the update observer registered by [`EditSession::set_update_observer`].
    pub fn clear_update_observer(&mut self) {
        self.update_observer = None;
    }

    /// Starts queueing committed transactions for
    /// [`EditSession::drain_update_event`] instead of pushing them through a
    /// callback. Idempotent while observation is running.
    pub fn start_update_event_observation(&mut self) -> Result<(), JsValue> {
        self.core
            .start_update_event_observation()
            .map_err(JsValue::from)
    }

    /// Pops the oldest queued transaction, in arrival order. Byte 0 is `1`
    /// when the transaction had no local origin and `0` otherwise; the rest is
    /// the v1 update. Empty when the queue is drained or observation never
    /// started.
    pub fn drain_update_event(&self) -> Vec<u8> {
        self.core.drain_update_event()
    }

    /// Stops queueing and discards anything not yet drained.
    pub fn clear_update_event_observation(&mut self) {
        self.core.clear_update_event_observation()
    }

    // -- local input state (undo + awareness selection) --

    /// Starts local-origin undo tracking across every story. Call this after
    /// import or seeding but before the first edit, so the initial document is
    /// not an undo step; later calls keep the history.
    pub fn track_undo(&self) {
        self.core.track_undo()
    }

    /// Closes the current undo capture without adding an empty step.
    pub fn add_undo_boundary(&self) {
        self.core.add_undo_boundary()
    }

    /// Changes grouping policy while retaining undo and redo history.
    pub fn set_undo_capture_mode(&self, mode: &str) -> Result<(), JsValue> {
        self.core.set_undo_capture_mode(mode).map_err(JsValue::from)
    }

    /// Pins the input span used to measure the next capture gap.
    pub fn set_undo_input_time(&self, first: f64, last: f64) {
        self.core.set_undo_input_time(first, last)
    }

    /// Clears the pinned input span.
    pub fn clear_undo_input_time(&self) {
        self.core.clear_undo_input_time()
    }

    /// Current undo grouping policy.
    pub fn undo_capture_mode(&self) -> String {
        self.core.undo_capture_mode()
    }

    /// Selects a story, closing capture unless manual grouping is selected.
    pub fn select_story(&self, story: &str) {
        self.core.select_story(story)
    }

    /// Reverts the latest local-origin step and reports whether anything was
    /// reverted. Remote and system transactions are excluded by the manager's
    /// tracked-origin policy; `false` before tracking starts.
    pub fn undo(&self) -> bool {
        self.core.undo()
    }

    /// Reapplies the latest locally undone transaction and reports whether
    /// anything was reapplied.
    pub fn redo(&self) -> bool {
        self.core.redo()
    }

    /// Whether [`EditSession::undo`] would revert something.
    pub fn can_undo(&self) -> bool {
        self.core.can_undo()
    }

    /// Whether [`EditSession::redo`] would reapply something.
    pub fn can_redo(&self) -> bool {
        self.core.can_redo()
    }

    /// Stories changed by the latest undo or redo, sorted.
    pub fn history_stories(&self) -> Vec<String> {
        self.core.history_stories()
    }

    /// Stores this peer's anchor and head as sticky positions, replacing any
    /// previous selection. Both endpoints must lie in `story`. The positions
    /// live outside the yrs document, so they are never serialized as content
    /// or carried in an update. `Assoc::After` makes a collapsed caret advance
    /// with text inserted at it.
    #[allow(clippy::too_many_arguments)]
    pub fn set_selection(
        &self,
        story: &str,
        anchor_para: &str,
        anchor_offset: u32,
        head_para: &str,
        head_offset: u32,
    ) -> Result<(), JsValue> {
        self.core
            .set_selection(story, anchor_para, anchor_offset, head_para, head_offset)
            .map_err(JsValue::from)
    }

    /// Encodes one Loc as opaque sticky-position bytes that keep pointing at
    /// the same content as the story is edited. Publish them over an awareness
    /// transport and resolve them with
    /// [`EditSession::resolve_sticky_position`].
    pub fn encode_sticky_position(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
    ) -> Result<Vec<u8>, JsValue> {
        self.core
            .encode_sticky_position(story, para_id, offset)
            .map_err(JsValue::from)
    }

    /// Resolves sticky bytes from [`EditSession::encode_sticky_position`] back
    /// to `{"story","paraId","offset"}`. Errors when the bytes are malformed
    /// or the position no longer resolves in `story`.
    pub fn resolve_sticky_position(&self, story: &str, position: &[u8]) -> Result<String, JsValue> {
        self.core
            .resolve_sticky_position(story, position)
            .map_err(JsValue::from)
    }

    /// This peer's current selection as `{"anchor":{"story","paraId","offset"},
    /// "head":{…}}`, or `"null"` before [`EditSession::set_selection`] is
    /// called. Errors when an endpoint no longer resolves.
    pub fn selection(&self) -> Result<String, JsValue> {
        self.core.selection().map_err(JsValue::from)
    }

    /// This peer's selection in transportable form:
    /// `{"story","anchor":[byte, …],"head":[byte, …]}` with sticky-encoded
    /// endpoints, or `"null"` before a selection is set. A peer resolves it
    /// with [`EditSession::resolve_encoded_selection`].
    pub fn encoded_selection(&self) -> Result<String, JsValue> {
        self.core.encoded_selection().map_err(JsValue::from)
    }

    /// Resolves another peer's encoded selection against this replica's
    /// current state: `{"anchor":{"story","paraId","offset"},"head":{…}}`.
    /// Errors on malformed bytes or an endpoint that no longer resolves.
    pub fn resolve_encoded_selection(
        &self,
        story: &str,
        anchor: &[u8],
        head: &[u8],
    ) -> Result<String, JsValue> {
        self.core
            .resolve_encoded_selection(story, anchor, head)
            .map_err(JsValue::from)
    }

    /// Stores a rectangular anchor-cell to head-cell selection outside the yrs
    /// document. `range_json` is a [`TableRange`]:
    /// `{"anchor":{"story","tableIndex","row","column"},"head":{…}}`. The table
    /// embed is held by a sticky index and each endpoint by its cell story's
    /// stable identity, so the selection survives unrelated edits and follows
    /// inserted or deleted rows and columns. Errors when the two endpoints are
    /// not in the same table, or when either cell does not resolve.
    pub fn set_cell_selection(&self, range_json: &str) -> Result<(), JsValue> {
        self.core
            .set_cell_selection(range_json)
            .map_err(JsValue::from)
    }

    /// The current local cell selection as a [`TableRange`], or `"null"`
    /// before one is set. An endpoint whose cell was deleted clamps to a
    /// surviving nearby cell. Errors when the table itself no longer resolves.
    pub fn cell_selection(&self) -> Result<String, JsValue> {
        self.core.cell_selection().map_err(JsValue::from)
    }

    pub fn clear_cell_selection(&self) -> Option<String> {
        self.core.clear_cell_selection()
    }

    pub fn part_metadata_json(&self) -> String {
        self.core.part_metadata_json()
    }

    pub fn set_part_metadata(&self, json: &str, remove_json: &str) -> Result<bool, JsValue> {
        self.core
            .set_part_metadata(json, remove_json)
            .map_err(JsValue::from)
    }

    pub fn document_properties_json(&self) -> String {
        self.core.document_properties_json()
    }

    pub fn set_document_property(&self, key: &str, json: &str) -> Result<bool, JsValue> {
        self.core
            .set_document_property(key, json)
            .map_err(JsValue::from)
    }

    pub fn list_sections_json(&self) -> Result<String, JsValue> {
        self.core.list_sections_json().map_err(JsValue::from)
    }

    pub fn section_margins_json(&self) -> String {
        self.core.section_margins_json()
    }

    pub fn set_section_margins(&self, section: &str, json: &str) -> Result<bool, JsValue> {
        self.core
            .set_section_margins(section, json)
            .map_err(JsValue::from)
    }

    /// Adds a story holding one paragraph with `initial_text` (which must not
    /// contain paragraph breaks), `p_style` and `alignment`. Receipt:
    /// `{"paraId"}` — the paragraph ending at the story's pilcrow. Errors when
    /// the story id already exists.
    pub fn create_story(
        &self,
        story_id: &str,
        initial_text: &str,
        p_style: &str,
        alignment: &str,
    ) -> Result<String, JsValue> {
        self.core
            .create_story(story_id, initial_text, p_style, alignment)
            .map_err(JsValue::from)
    }

    /// Removes one complete story and its content. Errors on an unknown story.
    pub fn delete_story(&self, story_id: &str) -> Result<(), JsValue> {
        self.core.delete_story(story_id).map_err(JsValue::from)
    }

    // -- native table ops --
    //
    // Tables are addressed in the resolved rectangular grid, not by OOXML row
    // and cell order: `TableLocator` is `{"story","tableIndex"}` with the
    // zero-based ordinal of the table embed in its story, `CellLoc` adds
    // `{"row","column"}`, and `TableRange` is `{"anchor":CellLoc,"head":
    // CellLoc}` covering the rectangle they span. Each cell's content lives in
    // its own story, so structural edits create and destroy stories.
    //
    // Every op below returns the same receipt:
    // `{"table":TableLocator,"rows","columns","createdStoryIds":[string, …],
    // "deletedStoryIds":[string, …],"changedStoryIds":[string, …],
    // "newParaIds":[string, …],"deletedTable":bool,"revisionIds":[string, …]}`
    // — `rows`/`columns` are the grid AFTER the op, the story lists let the
    // host follow cell content in and out of existence and invalidate what
    // the op rewrote, and `revisionIds` is non-empty only for the ops
    // that accept a suggesting author. All of them additionally error on an
    // unknown table or a cell outside the grid.

    /// Inserts a `rows` x `columns` table at `(story, para_id, offset)`,
    /// creating one story per cell. Errors when either dimension is zero.
    #[allow(clippy::too_many_arguments)]
    pub fn insert_table(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        rows: u32,
        columns: u32,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        self.core
            .insert_table(
                story,
                para_id,
                offset,
                rows,
                columns,
                author_name,
                author_date,
            )
            .map_err(JsValue::from)
    }

    /// Inserts a clipboard table fragment at `(story, para_id, offset)`.
    /// `fragment_json` is a [`DocxTableFragment`] (`packages/docx/src/yrs/clipboardFragment.ts`).
    /// Returns the [`TableFragmentReceipt`] JSON, with `after` carrying the
    /// caret right after the inserted table embed.
    pub fn insert_table_fragment(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        fragment_json: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        self.core
            .insert_table_fragment(
                story,
                para_id,
                offset,
                fragment_json,
                author_name,
                author_date,
            )
            .map_err(JsValue::from)
    }

    pub fn validate_table_fragment(&self, fragment_json: &str) -> Result<(), JsValue> {
        self.core
            .validate_table_fragment(fragment_json)
            .map_err(JsValue::from)
    }

    /// Inserts a row above (`after = false`) or below (`after = true`) the
    /// cell `at_json` names. `at_json` is a [`CellLoc`].
    pub fn insert_row(
        &self,
        at_json: &str,
        after: bool,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        self.core
            .insert_row(at_json, after, author_name, author_date)
            .map_err(JsValue::from)
    }

    /// Inserts a column left (`after = false`) or right (`after = true`) of
    /// the cell `at_json` ([`CellLoc`]) names. Always a plain local edit.
    pub fn insert_column(&self, at_json: &str, after: bool) -> Result<String, JsValue> {
        self.core
            .insert_column(at_json, after)
            .map_err(JsValue::from)
    }

    /// Deletes every row the [`TableRange`] `range_json` covers. In suggesting
    /// mode the rows are marked `trDel` instead of removed.
    pub fn delete_row(
        &self,
        range_json: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        self.core
            .delete_row(range_json, author_name, author_date)
            .map_err(JsValue::from)
    }

    /// Deletes every column the [`TableRange`] `range_json` covers, along with
    /// the stories of the cells removed. Always a plain local edit.
    pub fn delete_column(&self, range_json: &str) -> Result<String, JsValue> {
        self.core.delete_column(range_json).map_err(JsValue::from)
    }

    /// Removes the table `table_json` ([`TableLocator`]) names plus every one
    /// of its reachable cell stories. Always a plain local edit; the receipt
    /// has `deletedTable: true`.
    pub fn delete_table(&self, table_json: &str) -> Result<String, JsValue> {
        self.core.delete_table(table_json).map_err(JsValue::from)
    }

    /// Merges the rectangle the [`TableRange`] `range_json` covers into its
    /// top-left cell, whose story survives; the other cells' stories are
    /// deleted. Always a plain local edit. Errors when the range covers fewer
    /// than two cells or has no top-left anchor cell.
    pub fn merge_cells(&self, range_json: &str) -> Result<String, JsValue> {
        self.core.merge_cells(range_json).map_err(JsValue::from)
    }

    /// Splits the cell `at_json` ([`CellLoc`]) covers into a `rows` x
    /// `columns` grid; the original story stays in the top-left slot and every
    /// other slot gets a fresh one-paragraph cell story. Omitting both
    /// dimensions unmerges the cell into the slots it already covers, which
    /// then requires a merged cell. Always a plain local edit. Errors on a
    /// zero dimension or a grid smaller than the cell already covers.
    pub fn split_cell(
        &self,
        at_json: &str,
        rows: Option<u32>,
        columns: Option<u32>,
    ) -> Result<String, JsValue> {
        self.core
            .split_cell(at_json, rows, columns)
            .map_err(JsValue::from)
    }

    /// Sets every selected cell's background to `color` (hex, with or without
    /// a leading `#`), or clears it when `color` is absent. `range_json` is a
    /// [`TableRange`]. Always a plain local edit.
    pub fn set_cell_shading(
        &self,
        range_json: &str,
        color: Option<String>,
    ) -> Result<String, JsValue> {
        self.core
            .set_cell_shading(range_json, color)
            .map_err(JsValue::from)
    }

    /// Merges the JSON object `patch_json` into every selected cell's `tcPr`;
    /// a `null` value removes that key. `range_json` is a [`TableRange`].
    /// Always a plain local edit. Errors when the patch touches `rowspan`,
    /// `colspan`, `gridSpan` or `vMerge`, which only merge and split may set.
    pub fn set_cell_text_format(
        &self,
        range_json: &str,
        patch_json: &str,
    ) -> Result<String, JsValue> {
        self.core
            .set_cell_text_format(range_json, patch_json)
            .map_err(JsValue::from)
    }

    /// Merges the sides of the JSON object `borders_json` into every selected
    /// cell's `tcPr.borders`; `insideH`/`insideV` resolve to the physical edges
    /// interior to the selection. `range_json` is a [`TableRange`]. Always a
    /// plain local edit.
    pub fn set_cell_borders(
        &self,
        range_json: &str,
        borders_json: &str,
    ) -> Result<String, JsValue> {
        self.core
            .set_cell_borders(range_json, borders_json)
            .map_err(JsValue::from)
    }

    /// Sets the grid width, in twips, of the column holding the cell `at_json`
    /// ([`CellLoc`]) names. Always a plain local edit. Errors unless
    /// `width_twips` is finite and positive.
    pub fn set_column_width(&self, at_json: &str, width_twips: f64) -> Result<String, JsValue> {
        self.core
            .set_column_width(at_json, width_twips)
            .map_err(JsValue::from)
    }

    /// Sets the preferred width, in twips, of the table `table_json`
    /// ([`TableLocator`]) names. Always a plain local edit. Errors unless
    /// `width_twips` is finite and positive.
    pub fn set_table_width(&self, table_json: &str, width_twips: f64) -> Result<String, JsValue> {
        self.core
            .set_table_width(table_json, width_twips)
            .map_err(JsValue::from)
    }

    pub fn set_table_properties(
        &self,
        table_json: &str,
        patch_json: &str,
    ) -> Result<String, JsValue> {
        self.core
            .set_table_properties(table_json, patch_json)
            .map_err(JsValue::from)
    }

    /// Inserts `text` at `(story, para_id, offset)`. It must contain no
    /// paragraph or line breaks, and it inherits the formatting at the
    /// insertion point. Receipt: `{"revisionId": string|null, "range"}` —
    /// the id is non-null in suggesting mode, where the text is stamped `ins`
    /// and coalesces into an adjacent insertion by the same author rather than
    /// opening a second revision; the range is where the text landed.
    pub fn insert_text(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        text: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        self.core
            .insert_text(story, para_id, offset, text, author_name, author_date)
            .map_err(JsValue::from)
    }

    /// A plain local text edit that the resident lowering records, so the next
    /// [`Self::layout_document_with_regions_after_edit`] re-lowers and
    /// re-measures only its paragraph: inserts `text` at the collapsed range,
    /// or deletes `[start, end)` without it. The document changes exactly as
    /// through [`Self::insert_text`] or [`Self::delete_range`] without an
    /// author, and the receipt is theirs. An empty insertion, an edit that
    /// spans paragraphs, or one whose paragraph resident input is not ready
    /// for goes through them instead and records nothing for the relayout.
    pub fn edit_text_resident(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        text: Option<String>,
    ) -> Result<String, JsValue> {
        self.core
            .edit_text_resident(story, start_para, start_offset, end_para, end_offset, text)
            .map_err(JsValue::from)
    }

    /// Deletes `[start, end)`. Because a range crossing a paragraph boundary
    /// includes the boundary pilcrow, a plain delete also merges those
    /// paragraphs. Suggesting mode removes nothing and stamps the content
    /// `del` instead. Receipt: `{"revisionId": string|null, "range"}`, the
    /// range being what the delete left.
    #[allow(clippy::too_many_arguments)]
    pub fn delete_range(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        self.core
            .delete_range(
                story,
                start_para,
                start_offset,
                end_para,
                end_offset,
                author_name,
                author_date,
            )
            .map_err(JsValue::from)
    }

    /// Replaces `[start, end)` with `text` in one transaction. The inserted
    /// text adopts the first replaced unit's formatting; in suggesting mode
    /// the deletion and the insertion share one revision id. Receipt:
    /// `{"revisionId": string|null, "range": {"story", "start": {"paraId", "offset"}, "end": {"paraId", "offset"}}}`.
    #[allow(clippy::too_many_arguments)]
    pub fn replace_range(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        text: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        self.core
            .replace_range(
                story,
                start_para,
                start_offset,
                end_para,
                end_offset,
                text,
                author_name,
                author_date,
            )
            .map_err(JsValue::from)
    }

    pub fn insert_line_break(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        self.core
            .insert_line_break(story, para_id, offset, author_name, author_date)
            .map_err(JsValue::from)
    }

    /// Splits a paragraph at `(story, para_id, offset)` by inserting one
    /// pilcrow. The FIRST half keeps the original paraId and the second is
    /// re-minted. A split at the paragraph end leaves the empty second half
    /// with only the inherited property subset; a mid-paragraph split keeps
    /// its properties. Paragraph borders are cleared either way. Suggesting
    /// mode stamps the new pilcrow `ins` and `pPrIns`. Receipt:
    /// `{"firstParaId","secondParaId","revisionId": string|null}`.
    pub fn split_paragraph(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        self.core
            .split_paragraph(story, para_id, offset, author_name, author_date)
            .map_err(JsValue::from)
    }

    /// Merges `para_id` with the FOLLOWING paragraph by deleting (plain) or
    /// `del`- and `pPrDel`-marking (suggesting) its pilcrow. On a plain merge
    /// the survivor adopts the deleted mark's properties and paraId, so the
    /// earlier paragraph's identity wins. Receipt:
    /// `{"revisionId": string|null}`. Errors on the story's final paragraph,
    /// which has no following paragraph to merge with.
    pub fn merge_paragraphs(
        &self,
        story: &str,
        para_id: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        self.core
            .merge_paragraphs(story, para_id, author_name, author_date)
            .map_err(JsValue::from)
    }

    /// Applies one run mark over `[start, end)`. `mark_json`:
    /// `{"type":"bold"|"italic"|"underline"|"strike"|"superscript"|"subscript"} |
    /// {"type":"fontFamily"|"color","value":string} |
    /// {"type":"fontSize","value":number}`. The six boolean types TOGGLE —
    /// they turn on unless the whole range already carries the mark — while
    /// font family, size and color SET. Formatting is always a plain local
    /// edit, never a tracked change, so there is no receipt. Errors on an
    /// unknown `"type"` or a missing/mistyped `"value"`.
    #[allow(clippy::too_many_arguments)]
    pub fn toggle_mark(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        mark_json: &str,
    ) -> Result<(), JsValue> {
        self.core
            .toggle_mark(
                story,
                start_para,
                start_offset,
                end_para,
                end_offset,
                mark_json,
            )
            .map_err(JsValue::from)
    }

    /// Applies a set-valued, tri-state inline formatting delta over
    /// `[start, end)` in one transaction. An omitted key keeps the current
    /// value, `null` clears it, and any other value sets it. `delta_json`:
    ///
    /// ```json
    /// {
    ///   "bold": true, "italic": true,
    ///   "underline": true | {"style"?: string, "color"?: string},
    ///   "strike": true | {"double"?: boolean},
    ///   "color": {"rgb": string} | {"themeColor": string},
    ///   "highlight": string,
    ///   "fontSize": 12,
    ///   "fontFamily": {"ascii": string, "hAnsi"?: string},
    ///   "other": {"anyAttr": value}
    /// }
    /// ```
    ///
    /// `false` clears `bold`, `italic`, `underline` and `strike`. `color`
    /// takes exactly one of `rgb` or `themeColor`. `other` writes arbitrary
    /// run attributes, with `null` removing one. Always a plain local edit, so
    /// there is no receipt. Errors when a key carries a type not listed here.
    #[allow(clippy::too_many_arguments)]
    pub fn format_range(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        delta_json: &str,
    ) -> Result<(), JsValue> {
        self.core
            .format_range(
                story,
                start_para,
                start_offset,
                end_para,
                end_offset,
                delta_json,
            )
            .map_err(JsValue::from)
    }

    /// Sets or clears the hyperlink attribute over `[start, end)`.
    /// `hyperlink_json` is `{"href", "tooltip"?, "rId"?}` or `null` to unlink.
    /// The attribute is protected: ordinary formatting ops cannot write or
    /// erase it, only this one. Errors when the JSON is neither an object nor
    /// `null`.
    #[allow(clippy::too_many_arguments)]
    pub fn set_hyperlink(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        hyperlink_json: &str,
    ) -> Result<(), JsValue> {
        self.core
            .set_hyperlink(
                story,
                start_para,
                start_offset,
                end_para,
                end_offset,
                hyperlink_json,
            )
            .map_err(JsValue::from)
    }

    /// Clears every direct run-formatting attribute over `[start, end)`.
    /// Protected attributes — hyperlinks and tracked-change stamps — survive,
    /// as do paragraph properties.
    #[allow(clippy::too_many_arguments)]
    pub fn clear_formatting(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
    ) -> Result<(), JsValue> {
        self.core
            .clear_formatting(story, start_para, start_offset, end_para, end_offset)
            .map_err(JsValue::from)
    }

    /// The style a paragraph break at `(story, para_id, offset)` gives the new paragraph: the
    /// paragraph style's `w:next` when the break falls at the paragraph end and `w:next` names
    /// another paragraph style. `None` otherwise, and for a session without a source package.
    pub fn next_paragraph_style(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
    ) -> Result<Option<String>, JsValue> {
        self.core
            .next_paragraph_style(story, para_id, offset)
            .map_err(JsValue::from)
    }

    /// Applies a paragraph style and rebases inherited character formatting, preserving direct
    /// formatting. Without a source package, only `pStyle` changes.
    #[allow(clippy::too_many_arguments)]
    pub fn apply_paragraph_style(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        style_id: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<(), JsValue> {
        self.core
            .apply_paragraph_style(
                story,
                start_para,
                start_offset,
                end_para,
                end_offset,
                style_id,
                author_name,
                author_date,
            )
            .map_err(JsValue::from)
    }

    /// Applies a tri-state paragraph-property delta to every paragraph
    /// intersecting `[start, end)` in one transaction. An omitted key keeps
    /// the current value and `null` clears it. `attrs_json` recognises
    /// `alignment`, `lineSpacing`, `lineSpacingRule`, `spaceBefore`,
    /// `spaceAfter`, `indentLeft`, `indentRight`, `indentFirstLine`,
    /// `hangingIndent`, `bidi`, `tabs`
    /// (`[{"position":number,"alignment":string,"leader"?:string}, …]`) and
    /// `defaultTextFormatting` (an object of run defaults). Any other key —
    /// whether written at the top level or nested under `"other"` — is stored
    /// as an opaque paragraph property. Spacing and indents are authored OOXML
    /// units (twips, line-spacing units), never pixels. In suggesting mode a
    /// change is recorded as a `pPrChange` revision. Errors when a recognised
    /// key carries the wrong type, and when a key names schema-managed
    /// identity such as `paraId`.
    #[allow(clippy::too_many_arguments)]
    pub fn set_paragraph_attrs(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
        attrs_json: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<(), JsValue> {
        self.core
            .set_paragraph_attrs(
                story,
                start_para,
                start_offset,
                end_para,
                end_offset,
                attrs_json,
                author_name,
                author_date,
            )
            .map_err(JsValue::from)
    }

    /// Inserts one inline image embed at `(story, para_id, offset)`.
    /// The embed occupies one story unit. Receipt:
    /// `{"revisionId": string|null, "range"}`, the range being where the image
    /// landed. Errors when the payload is not an object.
    #[allow(clippy::too_many_arguments)]
    pub fn insert_image(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        payload_json: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String, JsValue> {
        self.core
            .insert_image(
                story,
                para_id,
                offset,
                payload_json,
                author_name,
                author_date,
            )
            .map_err(JsValue::from)
    }

    /// Moves one inline embed in one local transaction.
    #[allow(clippy::too_many_arguments)]
    pub fn move_embed(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        to_story: &str,
        to_para_id: &str,
        to_offset: u32,
    ) -> Result<String, JsValue> {
        self.core
            .move_embed(story, para_id, offset, to_story, to_para_id, to_offset)
            .map_err(JsValue::from)
    }

    /// Sets the authored `value` (any JSON) on the content-control embed
    /// carrying `embed_id`, searching every story. A plain- or rich-text
    /// control takes a string, which fills its content as one
    /// version-checked batch step instead. Errors when no embed has that id
    /// and when a fill is refused.
    pub fn set_content_control_value(
        &self,
        embed_id: &str,
        value_json: &str,
    ) -> Result<(), JsValue> {
        self.core
            .set_content_control_value(embed_id, value_json)
            .map_err(JsValue::from)
    }

    /// Sets the authored `value` on the content-control embed at
    /// `(story, para_id, offset)` — the way to reach a control with no
    /// authored `w:id` or tag, which
    /// [`EditSession::set_content_control_value`] cannot address. A text
    /// control's string value fills it as that method does. Errors when that
    /// position holds no embed.
    pub fn set_content_control_value_at(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        value_json: &str,
    ) -> Result<(), JsValue> {
        self.core
            .set_content_control_value_at(story, para_id, offset, value_json)
            .map_err(JsValue::from)
    }

    /// Removes the authored `value` from the content-control embed carrying
    /// `embed_id`, leaving the control and its content in place: it never
    /// erases a text control's text. Errors when no embed has that id.
    pub fn clear_content_control_value(&self, embed_id: &str) -> Result<(), JsValue> {
        self.core
            .clear_content_control_value(embed_id)
            .map_err(JsValue::from)
    }

    /// Writes geometry fields onto the image embed carrying `embed_id` in one
    /// transaction. Every key of the `geometry_json` object becomes a payload
    /// entry, with `null` clearing it; the entries of a nested `"other"`
    /// object are flattened into the payload alongside them. Errors when
    /// `geometry_json` or its `"other"` is not an object, and when no embed
    /// has that id.
    pub fn set_image_geometry(&self, embed_id: &str, geometry_json: &str) -> Result<(), JsValue> {
        self.core
            .set_image_geometry(embed_id, geometry_json)
            .map_err(JsValue::from)
    }

    /// [`EditSession::set_image_geometry`] for the image embed at
    /// `(story, para_id, offset)` — the way to reach one of several images
    /// sharing a relationship id, which the id variant resolves to the first.
    /// Errors when that position holds no image.
    pub fn set_image_geometry_at(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        geometry_json: &str,
    ) -> Result<(), JsValue> {
        self.core
            .set_image_geometry_at(story, para_id, offset, geometry_json)
            .map_err(JsValue::from)
    }

    /// Inserts a page-break embed at `(story, para_id, offset)`, occupying one
    /// story unit. Always a plain local edit.
    pub fn insert_page_break(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
    ) -> Result<(), JsValue> {
        self.core
            .insert_page_break(story, para_id, offset)
            .map_err(JsValue::from)
    }

    /// Inserts a section-break embed at `(story, para_id, offset)`.
    /// `break_type` must be `"nextPage"`, `"continuous"`, `"oddPage"` or
    /// `"evenPage"`; anything else errors. Always a plain local edit.
    pub fn insert_section_break(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        break_type: &str,
    ) -> Result<(), JsValue> {
        self.core
            .insert_section_break(story, para_id, offset, break_type)
            .map_err(JsValue::from)
    }

    /// Inserts a watermark embed at `(story, para_id, offset)`, its payload
    /// taken verbatim from the `watermark_json` object. Always a plain local
    /// edit. Errors when the JSON is not an object.
    pub fn insert_watermark(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        watermark_json: &str,
    ) -> Result<(), JsValue> {
        self.core
            .insert_watermark(story, para_id, offset, watermark_json)
            .map_err(JsValue::from)
    }

    /// Sets one paragraph property to any JSON value on `para_id`'s pilcrow,
    /// searching every story. Unlike
    /// [`EditSession::set_paragraph_attrs`] this writes a single key and never
    /// records a revision. Errors when the paragraph is unknown and when `key`
    /// names schema-managed identity (`paraId` or the embed discriminator).
    pub fn set_paragraph_attr(
        &self,
        para_id: &str,
        key: &str,
        value_json: &str,
    ) -> Result<(), JsValue> {
        self.core
            .set_paragraph_attr(para_id, key, value_json)
            .map_err(JsValue::from)
    }

    /// Adds a comment anchored to one or more ranges. `ranges_json`:
    /// `[{"story","startPara","startOffset","endPara","endOffset"}, …]`;
    /// `body_json` is any JSON value, stored as given. The anchors are sticky,
    /// so they follow the text they cover, and the comment lives outside the
    /// story rather than as an attribute on it. Receipt: `{"commentId"}` — a
    /// freshly minted id. Errors when `ranges_json` is not an array of
    /// well-formed ranges, or is empty.
    pub fn add_comment(
        &self,
        ranges_json: &str,
        author: &str,
        date: &str,
        body_json: &str,
    ) -> Result<String, JsValue> {
        self.core
            .add_comment(ranges_json, author, date, body_json)
            .map_err(JsValue::from)
    }

    /// Reanchors an existing comment without changing its metadata.
    pub fn set_comment_ranges(&self, comment_id: &str, ranges_json: &str) -> Result<(), JsValue> {
        self.core
            .set_comment_ranges(comment_id, ranges_json)
            .map_err(JsValue::from)
    }

    /// Initializes imported metadata, creating an unanchored comment when absent.
    pub fn seed_comment_metadata_json(
        &self,
        comment_id: &str,
        patch_json: &str,
    ) -> Result<(), JsValue> {
        self.core
            .seed_comment_metadata_json(comment_id, patch_json)
            .map_err(JsValue::from)
    }

    pub fn update_comment_metadata_json(
        &self,
        comment_id: &str,
        patch_json: &str,
    ) -> Result<(), JsValue> {
        self.core
            .update_comment_metadata_json(comment_id, patch_json)
            .map_err(JsValue::from)
    }

    /// Accepts tracked changes: pending insertions become plain content,
    /// pending deletions are carried out; `pPrIns` marks clear so the split
    /// stays, `pPrDel` marks join with the following paragraph, whose own
    /// properties survive.
    ///
    /// `target_json` is either `{"revisionId": string}`, resolving that one
    /// coalesced revision wherever it appears in any story, or
    /// `{"story","startPara","startOffset","endPara","endOffset"}`, resolving
    /// every tracked change overlapping that range regardless of revision id.
    /// Receipt: `{"revisionIds": [string, …]}` — the ids actually resolved, in
    /// resolution order and deduplicated. Resolving applies a revision rather
    /// than authoring one, so it never stamps a new revision and ignores
    /// suggesting mode. Errors on an empty range and on a `revisionId` that
    /// matches nothing.
    pub fn accept_change(&self, target_json: &str) -> Result<String, JsValue> {
        self.core.accept_change(target_json).map_err(JsValue::from)
    }

    /// Rejects tracked changes — the inverse of
    /// [`EditSession::accept_change`]: pending insertions roll back, pending
    /// deletions restore their text; `pPrIns` marks join back with the
    /// following paragraph, `pPrDel` marks clear so the split stays, and a
    /// `pPrChange` restores the paragraph's previous properties. Same target,
    /// receipt and error contract.
    pub fn reject_change(&self, target_json: &str) -> Result<String, JsValue> {
        self.core.reject_change(target_json).map_err(JsValue::from)
    }

    /// Resolves tracked changes by revision id outside undo history:
    /// `{"accept":[string, …],"reject":[string, …]}` -> `{"ok":true,"revisionIds":[string, …]}`,
    /// the ids resolved, or `{"ok":false,"version","failure"}`. An id that matches nothing is
    /// skipped. See [`EditingDoc::settle_revisions`].
    pub fn settle_revisions_json(&self, request: &str) -> Result<String, JsValue> {
        self.core
            .settle_revisions_json(request)
            .map_err(JsValue::from)
    }

    /// Applies a batch of raw story mutations in ONE transaction. Unlike every
    /// other op here these carry no user intent: indices are story-global
    /// UTF-16 units, not Locs, and nothing is inferred or stamped on the
    /// caller's behalf. `ops_json` is an array of:
    ///
    /// - `{"op":"insert","index","text"?,"attrs"?}`
    /// - `{"op":"delete","index","len"}`
    /// - `{"op":"format","index","len","attrs"?}`
    /// - `{"op":"insertEmbed","index","kind"?,"payload"?,"attrs"?}`
    /// - `{"op":"setEmbedAttr","index","key","value"?}`
    /// - `{"op":"setComment","id","ranges":[[start,end], …],"author"?,"date"?,"body"?}`
    /// - `{"op":"removeComment","id"}`
    ///
    /// Each op's `index`, and each `setComment` range, is read against the
    /// story state AFTER every preceding op in the batch. `attrs` and
    /// `payload` are JSON objects written verbatim, so tracked-change stamps
    /// travel inside `attrs`; comments are keyed by the given id and anchored
    /// sticky. Errors on an unknown `"op"`, a missing or negative `"index"`,
    /// or a missing required field, and leaves the story untouched — parsing
    /// completes before the transaction opens.
    pub fn apply_raw_ops(&self, story: &str, ops_json: &str) -> Result<(), JsValue> {
        self.core
            .apply_raw_ops(story, ops_json)
            .map_err(JsValue::from)
    }

    /// Applies seed raw operations with deterministic item ordering.
    pub fn apply_seed_raw_ops(&self, story: &str, ops_json: &str) -> Result<(), JsValue> {
        self.core
            .apply_seed_raw_ops(story, ops_json)
            .map_err(JsValue::from)
    }

    // -- version-checked host edits --
    //
    // Requests and results are JSON. Policy outcomes come back as
    // `{"ok":true,…}` or `{"ok":false,"version","failure":{"code","message",
    // "stepIndex"?,"conflictingStepIndex"?,"target"?}}`; malformed JSON,
    // unknown tags, missing fields and wrong types throw instead.

    /// The session-scoped version token of the committed document state. It
    /// changes on every committed change, local or remote, and when the
    /// document or its retained source is replaced.
    pub fn version(&self) -> String {
        self.core.version()
    }

    /// Versioned paragraph texts:
    /// `{"story"?,"paraIds"?,"view":"accepted"|"original"}` ->
    /// `{"ok":true,"version","view","paragraphs":[{"story","paraId","text",
    /// "styleId"?,"atoms":[{"offset","kind"}]}]}`. Each inline atom occupies one
    /// U+FFFC in `text`.
    pub fn read_paragraphs_json(&self, request: &str) -> Result<String, JsValue> {
        self.core
            .read_paragraphs_json(request)
            .map_err(JsValue::from)
    }

    /// Paragraph texts of selected stories in one read:
    /// `{"stories"?:"all"|[kind|id…],"byRoot"?,"view","expectVersion"?}` -> `{"ok":true,
    /// "version","view","truncated","stories":[{"story","ok":true,"paragraphs"}|{"story",
    /// "ok":false,"failure"}]}`. Listed entries keep their order; a kind expands in id order.
    pub fn read_stories_json(&self, request: &str) -> Result<String, JsValue> {
        self.core.read_stories_json(request).map_err(JsValue::from)
    }

    /// Every story: `{"ok":true,"version","stories":[{"story","kind","parent"?,"root","part"?,
    /// "uses"?:[{"sectionIndex","variant","shown"?}]}]}`, sorted by id. `shown` is present only
    /// while a complete layout of the current version is retained.
    pub fn list_stories_json(&self) -> Result<String, JsValue> {
        self.core.list_stories_json().map_err(JsValue::from)
    }

    /// Until the matching `end_shared_reads`, committed reads share story projections of each document state.
    pub fn begin_shared_reads(&self) {
        self.core.begin_shared_reads()
    }

    /// Ends a shared-read scope, dropping shared story projections when the last scope ends.
    pub fn end_shared_reads(&self) {
        self.core.end_shared_reads()
    }

    /// Exact, case-sensitive, paragraph-local search:
    /// `{"text","within","view","limit"?}` ->
    /// `{"ok":true,"version","matches":[{"text","range"}],"truncated"}`.
    pub fn find_text_json(&self, request: &str) -> Result<String, JsValue> {
        self.core.find_text_json(request).map_err(JsValue::from)
    }

    /// Runs every check of [`EditSession::apply_edits_json`], staging included,
    /// without changing anything: `{"ok":true,"baseVersion","wouldApply","previews"}`.
    pub fn validate_edits_json(&self, request: &str) -> Result<String, JsValue> {
        self.core
            .validate_edits_json(request)
            .map_err(JsValue::from)
    }

    /// Applies an edit batch all-or-nothing:
    /// `{"ok":true,"baseVersion","version","applied","source","changedStories",
    /// "receipts"}`. An applied batch commits one transaction; `history`
    /// `"separate"` makes it exactly one undo step.
    pub fn apply_edits_json(&self, request: &str) -> Result<String, JsValue> {
        self.core.apply_edits_json(request).map_err(JsValue::from)
    }

    /// Structured export of the committed document state:
    /// `{"revisionView","stories"?,"includeFormatting"?,"maxBlocks"?,"maxBytes"?}` ->
    /// `{"ok":true,"version","content"}` or `{"ok":false,"version","failure"}`. Anchors are scoped
    /// to the returned version. Reads only: nothing is committed, minted or published.
    pub fn export_structured_json(&self, options: &str) -> Result<String, JsValue> {
        self.core
            .export_structured_json(options)
            .map_err(JsValue::from)
    }

    /// [`EditSession::export_structured_json`] rendered as Markdown from the same read:
    /// `{"ok":true,"version","content":{"markdown","anchors","diagnostics","truncated"}}`.
    pub fn export_markdown_json(&self, options: &str) -> Result<String, JsValue> {
        self.core
            .export_markdown_json(options)
            .map_err(JsValue::from)
    }

    pub fn content_control_targets_json(&self) -> Result<String, JsValue> {
        self.core
            .content_control_targets_json()
            .map_err(JsValue::from)
    }

    pub fn repeat_content_control_at(
        &self,
        story: &str,
        index: u32,
        remove: bool,
    ) -> Result<bool, JsValue> {
        self.core
            .repeat_content_control_at(story, index, remove)
            .map_err(JsValue::from)
    }

    pub fn set_content_control_value_index(
        &self,
        story: &str,
        index: u32,
        json: &str,
    ) -> Result<(), JsValue> {
        self.core
            .set_content_control_value_index(story, index, json)
            .map_err(JsValue::from)
    }

    /// The content controls of the committed state:
    /// `{"stories"?,"maxControls"?,"maxBytes"?}` -> `{"ok":true,"version","content"}` or
    /// `{"ok":false,"version","failure"}`. Control ids and anchors are scoped to the returned
    /// version. Reads only.
    pub fn list_content_controls_json(&self, options: &str) -> Result<String, JsValue> {
        self.core
            .list_content_controls_json(options)
            .map_err(JsValue::from)
    }

    /// [`EditSession::list_content_controls_json`] keeping the controls that match `query`
    /// (`{"kind":"id","controlId"}`, `{"kind":"tag","tag"}`, `{"kind":"ooxmlId","ooxmlId"}` or
    /// `{"kind":"alias","alias"}`) exactly.
    pub fn find_content_controls_json(
        &self,
        query: &str,
        options: &str,
    ) -> Result<String, JsValue> {
        self.core
            .find_content_controls_json(query, options)
            .map_err(JsValue::from)
    }

    /// [`EditSession::export_structured_json`] with the page map of the retained region
    /// layout. `options` adds `"includeGeometry"?`, `"expectLayoutVersion"?`,
    /// `"maxFragments"?` and `"maxLayoutBytes"?`; the reply is
    /// `{"ok":true,"version","content":{"structured","layout"}}` or a refusal. An editor passes
    /// the region layout request it would lay the document out with now as `current_request`,
    /// and the layout must have been computed from the same inputs. Lays nothing out and changes
    /// nothing.
    pub fn export_structured_with_pages_json(
        &self,
        options: &str,
        current_request: Option<String>,
    ) -> Result<String, JsValue> {
        self.core
            .export_structured_with_pages_json(options, current_request)
            .map_err(JsValue::from)
    }

    /// Lays this private session out with its own fonts and exports it with pages as a
    /// snapshot. `fonts` holds the font files back to back, `font_lengths` their byte lengths;
    /// `request` is a region layout request whose font chains name fonts by their index. The
    /// reply is `{"ok":true,"content"}` or `{"ok":false,"failure"}`; a rejected font or an
    /// unusable request throws. The session's measurement fonts are left untouched.
    pub fn export_snapshot_with_private_fonts_json(
        &self,
        fonts: &[u8],
        font_lengths: &[u32],
        request: &str,
        options: &str,
    ) -> Result<String, JsValue> {
        self.core
            .export_snapshot_with_private_fonts_json(fonts, font_lengths, request, options)
            .map_err(JsValue::from)
    }

    /// Compares two DOCX packages into this session, which must hold no document: the original
    /// is opened here and the revised body text differences are applied as tracked changes,
    /// outside undo history.
    ///
    /// `options` is `{"author","date","granularity"?,"unsupported"?,"limits"?}`. Returns a final
    /// `{"ok":false,"diagnostics"}` or `{"ok":true,"noop":true,"changes":[],"diagnostics"}`, or
    /// `{"ok":true,"noop":false,"save"}` naming what a save of the applied changes may write:
    /// then pass the saved bytes to [`EditSession::finish_compared_docx_json`], or the save
    /// failure to [`EditSession::fail_compared_docx_json`]. Malformed options and internal
    /// failures throw.
    pub fn compare_docx_json(
        &self,
        original: &[u8],
        revised: &[u8],
        options: &str,
    ) -> Result<String, JsValue> {
        self.core
            .compare_docx_json(original, revised, options)
            .map_err(JsValue::from)
    }

    /// The final comparison result for `bytes`, the saved applied changes, verified against both
    /// inputs: `{"ok":true,"changes","diagnostics"}` or `{"ok":false,"diagnostics"}`. Throws
    /// when no comparison awaits its saved bytes.
    pub fn finish_compared_docx_json(&self, bytes: &[u8]) -> Result<String, JsValue> {
        self.core
            .finish_compared_docx_json(bytes)
            .map_err(JsValue::from)
    }

    /// The final comparison result when saving the applied changes failed with `message`.
    pub fn fail_compared_docx_json(&self, message: &str) -> Result<String, JsValue> {
        self.core
            .fail_compared_docx_json(message)
            .map_err(JsValue::from)
    }

    /// The headings of `story` in document order, classified as the structured export
    /// classifies them: `[{"paraId","heading":{"outlineLevel","source"}}]`.
    pub fn headings_json(&self, story: &str) -> Result<String, JsValue> {
        self.core.headings_json(story).map_err(JsValue::from)
    }

    /// Resolves a text target and formats it in one call (legacy agent
    /// helpers). `delta_json` is the [`EditSession::format_range`] delta.
    /// Returns `{"ok":true,"version"}` or a refusal.
    pub fn format_text_target_json(
        &self,
        target_json: &str,
        delta_json: &str,
    ) -> Result<String, JsValue> {
        self.core
            .format_text_target_json(target_json, delta_json)
            .map_err(JsValue::from)
    }

    /// Resolves a text target and anchors a side-map comment over it in one
    /// call (legacy agent helpers). `comment_json`:
    /// `{"id","author","date","body"?}`. Returns `{"ok":true,"version"}` or a
    /// refusal.
    pub fn comment_text_target_json(
        &self,
        target_json: &str,
        comment_json: &str,
    ) -> Result<String, JsValue> {
        self.core
            .comment_text_target_json(target_json, comment_json)
            .map_err(JsValue::from)
    }

    /// Accepted-view texts around a paragraph-keyed selection:
    /// `{"paraId","selectedText","paragraphText","before","after"}`, with
    /// `\n` between paragraphs and U+FFFC for each inline atom.
    #[allow(clippy::too_many_arguments)]
    pub fn selection_text_json(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
    ) -> Result<String, JsValue> {
        self.core
            .selection_text_json(story, start_para, start_offset, end_para, end_offset)
            .map_err(JsValue::from)
    }

    // -- read queries --
    //
    // Every query below is a pure snapshot of the document as it stands: it
    // mutates nothing, commits no transaction, and its JSON is fully
    // materialized before it returns.

    /// Aggregated toolbar and accessibility state over `[start, end)`:
    ///
    /// ```json
    /// {
    ///   "bold": true | false | "mixed", "italic": …, "underline": …, "strike": …,
    ///   "superscript": …, "subscript": …,
    ///   "fontFamily": string|null, "fontSize": number|null, "color": string|null,
    ///   "highlight": string|null,
    ///   "paraId": string, "styleId": string|null, "alignment": string|null,
    ///   "paragraphProperties": {…},
    ///   "hasSelection": bool, "isMultiParagraph": bool, "inTable": bool,
    ///   "isSingleEmbed": bool, "embedKind": string|null, "isImage": bool,
    ///   "inInsertion": bool, "inDeletion": bool
    /// }
    /// ```
    ///
    /// A toggle mark is `"mixed"` when the range disagrees about it; a value
    /// mark is `null` when it is absent OR not uniform, so the two cases are
    /// not distinguished. `isImage` is `embedKind == "image"`, and
    /// `inInsertion`/`inDeletion` report whether the range sits inside a
    /// pending tracked change.
    #[allow(clippy::too_many_arguments)]
    pub fn selection_context(
        &self,
        story: &str,
        start_para: &str,
        start_offset: u32,
        end_para: &str,
        end_offset: u32,
    ) -> Result<String, JsValue> {
        self.core
            .selection_context(story, start_para, start_offset, end_para, end_offset)
            .map_err(JsValue::from)
    }

    pub fn geometry_position_outline_json(&self, root: &str) -> Result<String, JsValue> {
        self.core
            .geometry_position_outline_json(root)
            .map_err(JsValue::from)
    }

    pub fn resident_geometry_json(&self, root: &str, request: &str) -> Result<String, JsValue> {
        self.core
            .resident_geometry_json(root, request)
            .map_err(JsValue::from)
    }

    pub fn resident_section_revision(&self) -> String {
        self.core.resident_section_revision()
    }

    pub fn proposal_revision_ranges_json(&self, ids_json: &str) -> Result<String, JsValue> {
        self.core
            .proposal_revision_ranges_json(ids_json)
            .map_err(JsValue::from)
    }

    /// Every pending tracked change across all stories, in deterministic
    /// story-then-position order:
    /// `[{"revisionId","author","date","kind","story","preview",
    /// "range":{"story","start":{"paraId","offset"},"end":{…}}}, …]`. `kind`
    /// is one of `"insertion"`, `"deletion"`, `"pPrIns"`, `"pPrDel"`,
    /// `"pPrChange"`, `"trIns"`, `"trDel"`, `"tableIns"` or `"tableDel"`.
    /// `preview` holds the first few characters of the affected text, and is
    /// empty for every structural kind, which covers no text.
    pub fn list_revisions(&self) -> Result<String, JsValue> {
        self.core.list_revisions().map_err(JsValue::from)
    }

    /// Author and date stamps for the requested revision ids.
    pub fn revision_stamps_json(&self, ids_json: &str) -> Result<String, JsValue> {
        self.core
            .revision_stamps_json(ids_json)
            .map_err(JsValue::from)
    }

    pub fn search_text(
        &self,
        query: &str,
        case_sensitive: bool,
        limit: Option<u32>,
    ) -> Result<String, JsValue> {
        self.core
            .search_text(query, case_sensitive, limit)
            .map_err(JsValue::from)
    }

    /// Whether the document has a story with this id.
    pub fn has_story(&self, story: &str) -> bool {
        self.core.has_story(story)
    }

    /// `{"revision","stories":[…]}`: the current story revision and the sorted
    /// ids of the stories created, edited, or deleted after revision `since`.
    pub fn stories_changed_since(&self, since: f64) -> String {
        self.core.stories_changed_since(since)
    }

    /// Every story id in the document, sorted so the order is stable across
    /// replicas.
    pub fn story_ids(&self) -> Vec<String> {
        self.core.story_ids()
    }

    /// Story length in UTF-16 units, every embed (pilcrows included) counting
    /// as one. Errors on an unknown story.
    pub fn story_len(&self, story: &str) -> Result<u32, JsValue> {
        self.core.story_len(story).map_err(JsValue::from)
    }

    /// The story's `canonical-stream-v1` FNV-1a checksum (see
    /// [`crate::canonical`]) as a DECIMAL STRING, because a u64 exceeds the
    /// JavaScript safe-integer range. Two stories with the same authored
    /// content share a checksum even when their paragraph and comment ids
    /// differ. Errors on an unknown story.
    pub fn story_checksum(&self, story: &str) -> Result<String, JsValue> {
        self.core.story_checksum(story).map_err(JsValue::from)
    }

    /// Lowers one story to a `LayoutBlock[]` JSON array — the block, run and
    /// table vocabulary the layout engine consumes. `env_json` supplies the
    /// document-level values lowering cannot read off the story:
    /// `{"themeColors":{slot: hex},"defaultTabStopTwips":number|null,
    /// "pageContentHeight":number|null,"numericIds":{yrsId: number},
    /// "tocStyleIds":[styleId],"showHiddenText":bool,
    /// "defaultParagraphStyleId":string,
    /// "paragraphStyleFontSizes":{styleId: pt},"docDefaultFontSize":number,
    /// "revisionPreview":{revisionId: "accepted"|"rejected"}}`, all
    /// optional. Errors when the story does not end in a pilcrow, holds a
    /// malformed table, references itself through a cell story, or contains an
    /// embed lowering does not support.
    pub fn yrs_blocks_for_story(&self, story: &str, env_json: &str) -> Result<String, JsValue> {
        self.core
            .yrs_blocks_for_story(story, env_json)
            .map_err(JsValue::from)
    }

    /// `[{"paraId","text","properties"}, …]` in document order. `text` is the
    /// paragraph's plain text without its pilcrow; `properties` is the
    /// pilcrow's authored property map (`pStyle`, `alignment` and whatever
    /// else has been set on it). Errors on an unknown story.
    pub fn paragraphs(&self, story: &str) -> Result<String, JsValue> {
        self.core.paragraphs(story).map_err(JsValue::from)
    }

    /// Compact paragraph-position projection built in one story traversal:
    /// `[{"paraId","length"}, …]` in document order, where `length` counts the
    /// UTF-16 text and inline embed units before that paragraph's pilcrow —
    /// exactly the `offset` domain of a Loc in it. One call replaces crossing
    /// the boundary once per paragraph. Errors on an unknown story.
    pub fn paragraph_spans(&self, story: &str) -> Result<String, JsValue> {
        self.core.paragraph_spans(story).map_err(JsValue::from)
    }

    /// The story as an ordered run of formatted segments — the same view
    /// lowering reads:
    ///
    /// - `{"kind":"text","text","attributes"}`
    /// - `{"kind":"pilcrow","paraId","properties","attributes"}`
    /// - `{"kind":"embed","embedKind","payload","attributes"}`
    ///
    /// A text segment covers one maximal run of identically formatted
    /// characters; each pilcrow and embed is its own segment worth one unit.
    /// `attributes` holds the segment's run marks together with any `ins`/`del`
    /// tracked-change stamps. Errors on an unknown story.
    pub fn story_segments(&self, story: &str) -> Result<String, JsValue> {
        self.core.story_segments(story).map_err(JsValue::from)
    }

    /// `story_segments` split after each pilcrow into units, as one hex digest
    /// per unit: `["digest", …]`. Equal digests mean equal segments.
    pub fn story_segment_unit_digests(&self, story: &str) -> Result<String, JsValue> {
        self.core
            .story_segment_unit_digests(story)
            .map_err(JsValue::from)
    }

    /// The segments of the listed units (indices into
    /// `story_segment_unit_digests`), each as `story_segments` gives them:
    /// `[[segment, …], …]`. Errors on an index past the last unit.
    pub fn story_segment_units(&self, story: &str, units: Vec<u32>) -> Result<String, JsValue> {
        self.core
            .story_segment_units(story, units)
            .map_err(JsValue::from)
    }

    /// The `payload` of the story's `table_index`-th table embed, as
    /// `story_segments` gives it, or `None` when the story has no such
    /// table. Errors on an unknown story.
    pub fn table_payload(&self, story: &str, table_index: u32) -> Result<Option<String>, JsValue> {
        self.core
            .table_payload(story, table_index)
            .map_err(JsValue::from)
    }

    /// `{"start","end"}` — the paragraph's span in story-global UTF-16 units.
    /// `end` is the index of its own pilcrow, so `end - start` is the
    /// paragraph length and the upper bound of a Loc `offset` in it. Errors
    /// when the paragraph is not in `story`.
    pub fn locate_paragraph(&self, story: &str, para_id: &str) -> Result<String, JsValue> {
        self.core
            .locate_paragraph(story, para_id)
            .map_err(JsValue::from)
    }

    /// The index of the first paragraph of `story` that carries `para_id`, in
    /// `paragraph_spans` order, or -1. Errors on an unknown story.
    pub fn paragraph_unit(&self, story: &str, para_id: &str) -> Result<i32, JsValue> {
        self.core
            .paragraph_unit(story, para_id)
            .map_err(JsValue::from)
    }

    /// Ids of the tracked changes of `story` whose range holds the caret at
    /// `para_id`/`offset`, in `list_revisions` order: `["id", …]`. The caret
    /// and both range ends resolve through their paragraph's span start as
    /// `locate_paragraph` gives it; a range whose paragraph does not resolve
    /// holds nothing. Reads only `story`, unlike `list_revisions`.
    pub fn revision_ids_at(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
    ) -> Result<String, JsValue> {
        self.core
            .revision_ids_at(story, para_id, offset)
            .map_err(JsValue::from)
    }

    /// How many paragraphs of `story` carry `para_id`: 0, 1, or 2 for two or
    /// more. Errors on an unknown story.
    pub fn paragraph_id_count(&self, story: &str, para_id: &str) -> Result<u32, JsValue> {
        self.core
            .paragraph_id_count(story, para_id)
            .map_err(JsValue::from)
    }

    /// Every comment the session holds, sorted by id:
    /// `[{"id","author","date","done","parentId","body"}, …]`, `parentId`
    /// null for a top-level comment and `body` the JSON value it was given.
    pub fn list_comments(&self) -> Result<String, JsValue> {
        self.core.list_comments().map_err(JsValue::from)
    }

    /// Where a comment's sticky anchors currently sit:
    /// `[{"story","start","end"}, …]`, one entry per anchored range, in
    /// story-global UTF-16 units. Errors on an unknown comment id and when an
    /// anchor no longer resolves.
    pub fn resolve_comment(&self, comment_id: &str) -> Result<String, JsValue> {
        self.core.resolve_comment(comment_id).map_err(JsValue::from)
    }
}

#[wasm_bindgen]
impl EditSession {
    /// Gives every paragraph that saves without a Word paragraph ID a fresh
    /// one and repairs duplicates; see [`EditingDoc::persist_paragraph_ids`].
    /// Receipt: `{"status":"applied","assignments":[{"paragraph",
    /// "replacedParaId","previousOoxmlParaId","ooxmlParaId","idOrigin",
    /// "persisted"}],"diagnostics":[…]}`, or `{"status":"refused","refusal"}`
    /// when nothing was changed.
    pub fn persist_paragraph_ids(&self) -> Result<String, JsValue> {
        self.core.persist_paragraph_ids().map_err(JsValue::from)
    }

    /// The Word paragraph ID each paragraph of `story` saves with, as a JSON
    /// array of strings or `null` in document order. Errors on an unknown story.
    pub fn story_paragraph_ids(&self, story: &str) -> Result<String, JsValue> {
        self.core.story_paragraph_ids(story).map_err(JsValue::from)
    }

    /// Every paragraph's identities: `{"sessionId","packageSha256",
    /// "paragraphs":[{"session","origin","ooxmlParaId","idOrigin",
    /// "persisted","source"}]}`, session paragraphs with stories sorted and
    /// in document order, then source paragraphs outside the stories, whose
    /// `session` is `null`. Reads only.
    pub fn paragraph_identities(&self) -> Result<String, JsValue> {
        self.core.paragraph_identities().map_err(JsValue::from)
    }

    /// Resolves a `session`, `source` or `persisted` anchor JSON against the
    /// current state: `{"status":"found","anchor"}`, `{"status":"missing"}`,
    /// `{"status":"ambiguous","candidates"}` or `{"status":"unsupported",
    /// "reason"}` with a `foreign-session`, `foreign-package` or
    /// `no-source-package` reason. A found source paragraph outside the
    /// session stories comes back as its source anchor. Errors on a malformed
    /// anchor.
    pub fn resolve_paragraph_anchor(&self, anchor_json: &str) -> Result<String, JsValue> {
        self.core
            .resolve_paragraph_anchor(anchor_json)
            .map_err(JsValue::from)
    }

    /// Caches save seed states without reseeding, within the paragraph and
    /// inflated package budgets.
    pub fn prime_save_seed_states(&self, paragraphs: u32, inflated: u32) -> bool {
        self.core.prime_save_seed_states(paragraphs, inflated)
    }

    /// The paragraph IDs a save applies, as the package writer's
    /// `paragraphIds` request field: `{"assignments":[{"part","ordinal",
    /// "paraId"}],"patchedParts":[{"part","paraIds":[[ordinal,"ID"]]}]}`,
    /// plus `"splicedParts":[{"part","sha256","paragraphs":[[ordinal,"key"]],
    /// "changed":[ordinal],"inserted":[{"before"|"after":ordinal}],"removed":[ordinal]}]`,
    /// whose session keys the caller resolves to the model paragraphs it marks
    /// with their ordinals.
    pub fn paragraph_save_plan(&self) -> Result<String, JsValue> {
        self.core.paragraph_save_plan().map_err(JsValue::from)
    }

    /// Reconciles and publishes the `[owner, paraId]` pairs a save captured,
    /// an owner being a session key or a source occurrence's
    /// `{partUri}#{ordinal}`; returns the stale pairs as the same JSON shape.
    /// See [`EditingDoc::record_saved_paragraph_ids`].
    pub fn record_saved_paragraph_ids(&self, saved_json: &str) -> Result<String, JsValue> {
        self.core
            .record_saved_paragraph_ids(saved_json)
            .map_err(JsValue::from)
    }

    /// The Word paragraph IDs a DOCX package holds, as JSON mapping each XML
    /// part URI to its paragraphs' IDs in document order, in canonical form.
    pub fn written_paragraph_ids(&self, bytes: &[u8]) -> Result<String, JsValue> {
        self.core
            .written_paragraph_ids(bytes)
            .map_err(JsValue::from)
    }

    /// Resolves what runs inherit under `styles_json`, a package's style definitions or `null`.
    pub fn run_inheritance(&self, styles_json: &str) -> Result<RunInheritance, JsValue> {
        let styles = self.core.run_inheritance_styles(styles_json)?;
        Ok(RunInheritance {
            inner: crate::seed::RunInheritance::new(Some(&styles)),
        })
    }
}

/// Structured export of DOCX bytes as a snapshot; `options` as for
/// [`EditSession::export_structured_json`]. No session is created.
#[wasm_bindgen]
pub fn export_docx_structured_json(bytes: &[u8], options: &str) -> Result<String, JsValue> {
    crate::session::free::export_docx_structured_json(bytes, options).map_err(js_err)
}

/// [`export_docx_structured_json`] rendered as Markdown.
#[wasm_bindgen]
pub fn export_docx_markdown_json(bytes: &[u8], options: &str) -> Result<String, JsValue> {
    crate::session::free::export_docx_markdown_json(bytes, options).map_err(js_err)
}

/// The content controls of DOCX bytes as a snapshot; `options` as for
/// [`EditSession::list_content_controls_json`]. No session is created.
#[wasm_bindgen]
pub fn list_docx_content_controls_json(bytes: &[u8], options: &str) -> Result<String, JsValue> {
    crate::session::free::list_docx_content_controls_json(bytes, options).map_err(js_err)
}

/// The content controls of DOCX bytes that match `query` exactly.
#[wasm_bindgen]
pub fn find_docx_content_controls_json(
    bytes: &[u8],
    query: &str,
    options: &str,
) -> Result<String, JsValue> {
    crate::session::free::find_docx_content_controls_json(bytes, query, options).map_err(js_err)
}

/// Renders structured content as Markdown: `content` is a schema-version-1 export and
/// `options` is `{"maxBytes"?}`.
#[wasm_bindgen]
pub fn render_docx_markdown_json(content: &str, options: &str) -> Result<String, JsValue> {
    crate::session::free::render_docx_markdown_json(content, options).map_err(js_err)
}

/// Renders a paged export as Markdown: `content` is `{"structured","layout"}` from a paged
/// export and `options` is `{"maxBytes"?,"pageMarkers"?}`.
#[wasm_bindgen]
pub fn render_docx_markdown_with_pages_json(
    content: &str,
    options: &str,
) -> Result<String, JsValue> {
    crate::session::free::render_docx_markdown_with_pages_json(content, options).map_err(js_err)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::presence::encode_sticky;
    use crate::session::args::{index_loc, loc_index};
    use crate::{
        EditCtx, EditRequest, EditingDoc, FormatPolicy, ParagraphRef, Position, RawOp,
        SegmentContent, StoryRange, TriState, UndoSession, story_ref,
    };
    use serde_json::Value;
    use yrs::{Any, Assoc, IndexedSequence, Map, ReadTxn, Transact};

    #[test]
    fn collapsed_command_positions_share_sticky_walks() {
        use crate::sticky_positions::{STICKY_INDEX_WALKS, STICKY_OFFSET_WALKS};

        let session = EditSession::new(77.0).unwrap();
        let doc = session.core.engine.doc();
        doc.create_story_with_paragraph_id("body", "p1", "abcdef", "Normal", "left")
            .unwrap();
        let txn = doc.yrs_doc().transact();
        let text = story_ref(&txn, "body").unwrap();
        let expected = text.sticky_index(&txn, 3, Assoc::After).unwrap();
        drop(txn);
        STICKY_INDEX_WALKS.with(|walks| walks.set(0));
        STICKY_OFFSET_WALKS.with(|walks| walks.set(0));
        session.set_selection("body", "p1", 3, "p1", 3).unwrap();
        for _ in 0..2 {
            let bytes = session.encode_sticky_position("body", "p1", 3).unwrap();
            assert_eq!(bytes, encode_sticky(&expected));
            let resolved: Value =
                serde_json::from_str(&session.resolve_sticky_position("body", &bytes).unwrap())
                    .unwrap();
            assert_eq!(
                resolved,
                json!({"story": "body", "paraId": "p1", "offset": 3})
            );
        }
        let selection = session.core.selection.borrow();
        let selection = selection.as_ref().unwrap();
        assert_eq!(selection.anchor, expected);
        assert_eq!(selection.head, expected);
        assert_eq!(
            doc.cached_sticky_offset("body", &selection.head).unwrap(),
            Some(3)
        );
        STICKY_INDEX_WALKS.with(|walks| assert_eq!(walks.get(), 1));
        STICKY_OFFSET_WALKS.with(|walks| assert_eq!(walks.get(), 1));
    }

    #[test]
    fn collapsed_input_positions_share_sticky_walks() {
        use crate::sticky_positions::{STICKY_INDEX_WALKS, STICKY_OFFSET_WALKS};

        for repeats in [1, 32, 256] {
            for reads in [1, 8] {
                let session = EditSession::new(77.0).unwrap();
                let doc = session.core.engine.doc();
                let initial = "abcdef".repeat(repeats);
                let offset = initial.len() as u32 / 2;
                doc.create_story_with_paragraph_id("body", "p1", &initial, "Normal", "left")
                    .unwrap();
                let font_id = session
                    .register_measure_font(include_bytes!(
                        "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
                    ))
                    .unwrap();
                let request = json!({
                    "bodyStory": "body",
                    "regions": {"sections": [{
                        "sectionId": "main",
                        "properties": {
                            "pageWidth": 4320,
                            "pageHeight": 2880,
                            "marginTop": 300,
                            "marginRight": 300,
                            "marginBottom": 300,
                            "marginLeft": 300
                        }
                    }]},
                    "measurement": {
                        "fontChains": {"calibri|0|0": [font_id]},
                        "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
                        "authoritativeShaping": true
                    },
                    "renderEnv": {}
                });
                session
                    .layout_document_with_regions_json(&request.to_string())
                    .unwrap();
                session
                    .build_display_list_frame(
                        &json!({"fontChains": {"calibri|0|0": [font_id]}}).to_string(),
                        0.0,
                    )
                    .unwrap();
                session.core.engine.set_display_window(Some(0..1));
                session.set_windowed_incremental_builds(true);
                STICKY_INDEX_WALKS.with(|walks| walks.set(0));
                STICKY_OFFSET_WALKS.with(|walks| walks.set(0));
                session
                    .set_selection("body", "p1", offset, "p1", offset)
                    .unwrap();
                let bytes = session
                    .encode_sticky_position("body", "p1", offset)
                    .unwrap();
                let before: Value =
                    serde_json::from_str(&session.resolve_sticky_position("body", &bytes).unwrap())
                        .unwrap();
                assert_eq!(before["offset"], offset);

                session.apply_input("x", 1.0).unwrap();
                for _ in 0..reads {
                    let selection: Value =
                        serde_json::from_str(&session.selection().unwrap()).unwrap();
                    let peer: Value = serde_json::from_str(
                        &session
                            .resolve_encoded_selection("body", &bytes, &bytes)
                            .unwrap(),
                    )
                    .unwrap();
                    assert_eq!(selection, peer);
                    assert_eq!(selection["anchor"]["offset"], offset + 1);
                    assert_eq!(selection["head"], selection["anchor"]);
                    let snapshot: Value =
                        serde_json::from_str(&session.resident_caret_snapshot_json().unwrap())
                            .unwrap();
                    assert_eq!(snapshot["frameEpoch"], 2);
                    assert!(snapshot["caretRect"].is_object());
                }
                let mut expected = initial;
                expected.insert(offset as usize, 'x');
                assert_eq!(doc.paragraphs("body").unwrap()[0].text, expected);
                STICKY_INDEX_WALKS.with(|walks| assert_eq!(walks.get(), 1));
                STICKY_OFFSET_WALKS.with(|walks| assert_eq!(walks.get(), 2));
            }
        }
    }

    #[test]
    fn line_break_inserts_one_inline_unit_and_is_undoable() {
        let mut session = EditSession::new(76.0).unwrap();
        session.core.undo = UndoSession::with_clock(Arc::new(|| 0));
        session
            .core
            .engine
            .doc()
            .create_story_with_paragraph_id("body", "p1", "ab", "Normal", "left")
            .unwrap();
        session.track_undo();
        let receipt: Value = serde_json::from_str(
            &session
                .insert_line_break("body", "p1", 1, None, None)
                .unwrap(),
        )
        .unwrap();
        assert_eq!(receipt["range"]["start"]["offset"], 1);
        assert_eq!(receipt["range"]["end"]["offset"], 2);
        let segments = session.story_segments("body").unwrap();
        let parsed: Value = serde_json::from_str(&segments).unwrap();
        assert_eq!(parsed[0]["text"], "a");
        assert_eq!(parsed[1]["embedKind"], "break");
        assert_eq!(parsed[2]["text"], "b");
        assert_eq!(
            session.core.engine.doc().paragraphs("body").unwrap().len(),
            1
        );
        session.add_undo_boundary();
        assert!(session.undo());
        assert_eq!(
            session.core.engine.doc().paragraphs("body").unwrap()[0].text,
            "ab"
        );
        assert!(session.redo());
        assert_eq!(session.story_segments("body").unwrap(), segments);
    }

    #[test]
    fn paragraph_segments_build_only_requested_units() {
        let session = EditSession::new(75.0).unwrap();
        let doc = session.core.engine.doc();
        doc.create_story("body", "A😀B", "Normal", "left").unwrap();
        for _ in 0..80 {
            doc.split_paragraph(&EditCtx::local("", ""), Position::new("body", 0), None)
                .unwrap();
        }
        let requested = vec![80, 0, 80];
        doc.segment_index("body").unwrap();
        let assert_targeted = |rebuild: bool| {
            let all = doc.story_segment_units("body").unwrap();
            let expected: Vec<_> = requested
                .iter()
                .map(|&unit| {
                    crate::segment_json::segments_json(all[unit as usize].clone()).unwrap()
                })
                .collect();
            let work: usize = requested.iter().map(|&unit| all[unit as usize].len()).sum();
            crate::STORY_SEGMENTS_BUILT.with(|built| built.set(0));
            crate::STORY_INDEX_UNITS_BUILT.with(|built| built.set(0));
            let actual: Value = serde_json::from_str(
                &session
                    .story_segment_units("body", requested.clone())
                    .unwrap(),
            )
            .unwrap();
            assert_eq!(actual, json!(expected));
            crate::STORY_SEGMENTS_BUILT.with(|built| assert_eq!(built.get(), work));
            crate::STORY_INDEX_UNITS_BUILT.with(|built| {
                if rebuild {
                    assert!(built.get() >= 81);
                } else {
                    assert_eq!(built.get(), 0);
                }
            });
            assert!(
                doc.story_segment_units_at("body", &[81])
                    .unwrap_err()
                    .contains("no segment unit 81 in body")
            );
        };
        assert_targeted(false);
        let para_id = doc.paragraphs("body").unwrap()[80].para_id.clone();
        let before = doc.committed_epoch();
        let previous = doc.segment_index("body").unwrap();
        crate::STORY_INDEX_UNITS_BUILT.with(|built| built.set(0));
        session
            .insert_text("body", &para_id, 1, "!", None, None)
            .unwrap();
        doc.advance_indexes_after_text_insert("body", before, doc.committed_epoch(), 81, "!");
        crate::STORY_INDEX_UNITS_BUILT.with(|built| assert!(built.get() <= 3));
        let current = doc.segment_index("body").unwrap();
        assert_eq!(
            previous.segment_units().unwrap()[0].as_ptr(),
            current.segment_units().unwrap()[0].as_ptr()
        );
        assert_targeted(false);
        crate::STORY_INDEX_UNITS_BUILT.with(|built| built.set(0));
        session
            .delete_range("body", &para_id, 1, &para_id, 2, None, None)
            .unwrap();
        crate::STORY_INDEX_UNITS_BUILT.with(|built| assert!(built.get() <= 3));
        assert_targeted(false);
        session
            .format_range("body", &para_id, 0, &para_id, 1, r#"{"bold":true}"#)
            .unwrap();
        assert_targeted(true);
        crate::STORY_INDEX_UNITS_BUILT.with(|built| built.set(0));
        session
            .insert_text("body", &para_id, 1, "?", None, None)
            .unwrap();
        crate::STORY_INDEX_UNITS_BUILT.with(|built| assert!(built.get() <= 4));
        assert_targeted(false);
        crate::STORY_INDEX_UNITS_BUILT.with(|built| built.set(0));
        session
            .delete_range("body", &para_id, 2, &para_id, 4, None, None)
            .unwrap();
        crate::STORY_INDEX_UNITS_BUILT.with(|built| assert!(built.get() <= 4));
        assert_targeted(false);
        let remote = EditSession::new(76.0).unwrap();
        remote.load(&session.encode_state()).unwrap();
        remote
            .insert_text("body", &para_id, 0, "--", None, None)
            .unwrap();
        session.apply_update(&remote.encode_state()).unwrap();
        assert_targeted(true);
        crate::STORY_SEGMENTS_BUILT.with(|built| built.set(0));
        assert_eq!(
            session.story_segment_units("body", Vec::new()).unwrap(),
            "[]"
        );
        crate::STORY_SEGMENTS_BUILT.with(|built| assert_eq!(built.get(), 0));
    }

    #[test]
    fn selection_context_reads_only_selected_units_after_text_edits() {
        let session = EditSession::new(77.0).unwrap();
        let doc = session.core.engine.doc();
        doc.create_story("body", "Seed", "Normal", "left").unwrap();
        for _ in 0..80 {
            doc.split_paragraph(&EditCtx::local("", ""), Position::new("body", 0), None)
                .unwrap();
        }
        let para_id = doc.paragraphs("body").unwrap()[80].para_id.clone();
        for index in 0..40 {
            doc.create_story(format!("aux-{index}"), "Other", "Normal", "left")
                .unwrap();
        }
        let read_context = |story: &str, para_id: &str| {
            let (start, _) = doc
                .paragraph_index(story)
                .unwrap()
                .para_span(para_id)
                .unwrap();
            doc.selection_context(&StoryRange::new(story, start + 1, start + 1))
                .unwrap()
        };
        read_context("body", &para_id);
        let assert_targeted = || {
            crate::SNAPSHOT_CHUNKS_BUILT.with(|built| built.set(0));
            crate::TABLE_MEMBERSHIP_STORIES_SCANNED.with(|scanned| scanned.set(0));
            let context = read_context("body", &para_id);
            assert_eq!(context.para_id, para_id);
            assert!(!context.in_table);
            assert_eq!(context.bold, TriState::Off);
            crate::STORY_INDEX_UNITS_BUILT.with(|built| assert!(built.get() <= 3));
            crate::SNAPSHOT_CHUNKS_BUILT.with(|built| assert_eq!(built.get(), 2));
            crate::TABLE_MEMBERSHIP_STORIES_SCANNED.with(|scanned| assert_eq!(scanned.get(), 0));
        };
        crate::STORY_INDEX_UNITS_BUILT.with(|built| built.set(0));
        session
            .insert_text("body", &para_id, 1, "!", None, None)
            .unwrap();
        assert_targeted();
        crate::STORY_INDEX_UNITS_BUILT.with(|built| built.set(0));
        session
            .delete_range("body", &para_id, 1, &para_id, 2, None, None)
            .unwrap();
        assert_targeted();
        session
            .format_range("body", &para_id, 0, &para_id, 1, r#"{"bold":true}"#)
            .unwrap();
        crate::TABLE_MEMBERSHIP_STORIES_SCANNED.with(|scanned| scanned.set(0));
        let context = read_context("body", &para_id);
        assert_eq!(context.bold, TriState::On);
        crate::TABLE_MEMBERSHIP_STORIES_SCANNED.with(|scanned| assert_eq!(scanned.get(), 0));
        let cell = Any::Map(Arc::new(std::collections::HashMap::from([(
            "story".into(),
            Any::from("aux-0"),
        )])));
        let row = Any::Map(Arc::new(std::collections::HashMap::from([(
            "cells".into(),
            Any::Array(Arc::from(vec![cell])),
        )])));
        doc.apply_raw_ops(
            "body",
            vec![RawOp::InsertEmbed {
                index: 0,
                kind: "table".into(),
                payload: vec![("rows".into(), Any::Array(Arc::from(vec![row])))],
                attrs: yrs::types::Attrs::new(),
            }],
            &EditCtx::local("", ""),
        )
        .unwrap();
        crate::TABLE_MEMBERSHIP_STORIES_SCANNED.with(|scanned| scanned.set(0));
        let cell_id = doc.paragraphs("aux-0").unwrap()[0].para_id.clone();
        let context = read_context("aux-0", &cell_id);
        assert!(context.in_table);
        crate::TABLE_MEMBERSHIP_STORIES_SCANNED.with(|scanned| assert_eq!(scanned.get(), 41));
    }

    #[test]
    fn header_footer_alias_exports_seed_and_follow_legacy_updates() {
        let session = EditSession::new(41.0).unwrap();
        let before = session.core.engine.doc().encode_state_as_update_v1();
        session.set_header_footer_aliases("[]").unwrap();
        assert_eq!(
            session.core.engine.doc().encode_state_as_update_v1(),
            before
        );
        assert_eq!(session.header_footer_aliases_json(), "{}");
        let groups = json!([{
            "isHeader": true, "partPath": "word/header1.xml", "relationshipIds": ["rId7", "rId9"]
        }]);
        session
            .set_header_footer_aliases(&groups.to_string())
            .unwrap();
        session
            .core
            .engine
            .doc()
            .create_story("hf:rId7", "Header", "Normal", "left")
            .unwrap();
        assert_eq!(session.header_footer_aliases_json(), r#"{"rId9":"rId7"}"#);
        let peer = EditingDoc::new(42);
        peer.apply_update_v1(&session.core.engine.doc().encode_state_as_update_v1())
            .unwrap();
        peer.create_story("hf:rId9:t0:r0c0", "Legacy", "Normal", "left")
            .unwrap();
        session
            .core
            .engine
            .doc()
            .apply_update_v1(&peer.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(session.header_footer_aliases_json(), "{}");
    }

    #[test]
    fn docx_open_keeps_aliases_inactive() {
        let bytes = crate::seed::header_footer_alias_fixture::package(
            &[("rId7", "header1.xml"), ("rId9", "./header1.xml")],
            &[],
        );
        for seed in [true, false] {
            let session = EditSession::new(41.0).unwrap();
            let result: Value = serde_json::from_str(
                &session
                    .core
                    .open_docx(&bytes, seed, Some("aliases".to_owned()), None)
                    .unwrap(),
            )
            .unwrap();
            assert_eq!(session.header_footer_aliases_json(), "{}");
            let txn = session.core.engine.doc().yrs_doc().transact();
            assert!(
                !txn.get_map(crate::identity::SESSION)
                    .unwrap()
                    .contains_key(&txn, "hfAliases")
            );
            drop(txn);
            assert!(
                result["envelope"]["document"]["package"]
                    .get("headerFooterAliases")
                    .is_none()
            );
            if seed {
                assert!(session.core.engine.doc().paragraphs("hf:rId7").is_ok());
                assert!(session.core.engine.doc().paragraphs("hf:rId9").is_ok());
            }
        }
    }

    #[test]
    fn owned_full_open_keeps_input_allocation_and_matches_borrowed_open() {
        for seed_stories in [false, true] {
            let bytes = unidentified_docx("雪 &amp; café & 🦀");
            let borrowed = EditSession::new(74.0).unwrap();
            let expected = borrowed
                .core
                .open_docx(&bytes, seed_stories, Some("fixed".to_owned()), None)
                .unwrap();
            let owned = EditSession::new(74.0).unwrap();
            let pointer = bytes.as_ptr();
            let actual = owned
                .open_docx_owned(bytes, seed_stories, Some("fixed".to_owned()), None)
                .unwrap();
            assert_eq!(actual, expected);
            assert_eq!(owned.encode_state(), borrowed.encode_state());
            assert_eq!(owned.encode_state_vector(), borrowed.encode_state_vector());
            assert_eq!(
                owned.materialize_docx().unwrap(),
                borrowed.materialize_docx().unwrap()
            );
            assert_eq!(owned.media_sources_json(), borrowed.media_sources_json());
            assert_eq!(
                owned.core.docx_source.borrow().as_ref().unwrap().as_ptr(),
                pointer
            );
            assert_eq!(
                owned
                    .core
                    .engine
                    .doc()
                    .source_index()
                    .unwrap()
                    .bytes()
                    .as_ptr(),
                pointer
            );
        }
        let bytes = batch_docx();
        let borrowed = EditSession::new(74.0).unwrap();
        let expected = borrowed
            .core
            .seed_from_docx(&bytes, Some("fixed".to_owned()))
            .unwrap();
        let owned = EditSession::new(74.0).unwrap();
        let pointer = bytes.as_ptr();
        let actual = owned
            .seed_from_docx_owned(bytes, Some("fixed".to_owned()))
            .unwrap();
        assert_eq!(actual, expected);
        assert_eq!(owned.encode_state(), borrowed.encode_state());
        assert_eq!(owned.encode_state_vector(), borrowed.encode_state_vector());
        assert_eq!(
            owned.core.docx_source.borrow().as_ref().unwrap().as_ptr(),
            pointer
        );
    }

    #[test]
    fn owned_preview_open_matches_borrowed_open() {
        for blocks in [1, 2, 3] {
            let bytes = batch_docx();
            let borrowed = EditSession::new(74.0).unwrap();
            let expected = borrowed.core.open_docx_preview(&bytes, blocks).unwrap();
            let owned = EditSession::new(74.0).unwrap();
            let actual = owned.open_docx_preview_owned(bytes, blocks).unwrap();
            assert_eq!(actual, expected);
            assert_eq!(owned.encode_state(), borrowed.encode_state());
            assert_eq!(owned.encode_state_vector(), borrowed.encode_state_vector());
            assert_eq!(
                owned.materialize_docx().unwrap(),
                borrowed.materialize_docx().unwrap()
            );
        }
        let bytes = script_fonts_docx(
            r#"<w:p><w:r><w:t>First</w:t></w:r></w:p><w:sectPr><w:cols w:num="2"/></w:sectPr>"#,
            "",
        );
        let borrowed = EditSession::new(74.0).unwrap();
        let owned = EditSession::new(74.0).unwrap();
        let expected = borrowed.core.open_docx_preview(&bytes, 1).unwrap();
        assert!(expected.is_none());
        assert_eq!(owned.open_docx_preview_owned(bytes, 1).unwrap(), expected);
        assert_eq!(owned.encode_state(), borrowed.encode_state());
        assert_eq!(owned.encode_state_vector(), borrowed.encode_state_vector());
    }

    #[test]
    fn owned_weighted_preview_open_matches_borrowed_open() {
        let p = "<w:p><w:r><w:t>Cell</w:t></w:r></w:p>";
        let table = format!("<w:tbl><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>", p.repeat(10));
        let bytes = script_fonts_docx(&table.repeat(60), "");
        for paragraph_budget in [None, Some(256), Some(400)] {
            let borrowed = EditSession::new(74.0).unwrap();
            let expected = borrowed
                .core
                .open_docx_preview_with_budget(&bytes, 200, paragraph_budget)
                .unwrap();
            let owned = EditSession::new(74.0).unwrap();
            let actual = owned
                .open_docx_preview_with_budget_owned(bytes.clone(), 200, paragraph_budget)
                .unwrap();
            assert_eq!(actual, expected);
            assert_eq!(owned.encode_state(), borrowed.encode_state());
            assert_eq!(owned.encode_state_vector(), borrowed.encode_state_vector());
            assert_eq!(
                owned.materialize_docx().unwrap(),
                borrowed.materialize_docx().unwrap()
            );
        }
    }

    #[test]
    fn owned_open_matches_borrowed_validation_errors() {
        for bytes in [Vec::new(), b"not a zip".to_vec(), batch_docx()] {
            for digest in [None, Some("malformed")] {
                let borrowed = EditSession::new(74.0).unwrap();
                let owned = EditSession::new(74.0).unwrap();
                assert_eq!(
                    owned.core.open_docx_retaining_source(
                        bytes.clone().into(),
                        true,
                        Some("fixed"),
                        digest
                    ),
                    borrowed
                        .core
                        .open_docx_retaining(&bytes, true, Some("fixed"), digest)
                );
                assert_eq!(owned.encode_state(), borrowed.encode_state());
                assert_eq!(owned.encode_state_vector(), borrowed.encode_state_vector());
                assert_eq!(
                    owned
                        .core
                        .open_preview_source(b"not a zip".to_vec(), 1, None),
                    borrowed.core.open_preview(b"not a zip", 1)
                );
            }
        }
    }

    #[test]
    fn seeded_docx_retains_original_images_for_materialization_and_save() {
        let image_bytes = vec![1, 2, 3, 4];
        let source = ooxml_opc::rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec()),
            ("word/_rels/document.xml.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/original.png"/></Relationships>"#.to_vec()),
            ("word/document.xml".to_owned(), br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="original"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>"#.to_vec()),
            ("word/media/original.png".to_owned(), image_bytes.clone()),
        ])
        .unwrap();
        let expected = crate::seed::parse_docx_for_edit(&source).unwrap();
        assert!(!expected.document.package.media_entries.is_empty());
        let session = EditSession::new(7.0).unwrap();
        let host: Value =
            serde_json::from_str(&session.core.open_docx(&source, true, None, None).unwrap())
                .unwrap();
        assert_eq!(
            host["envelope"]["document"]["package"]["mediaEntries"],
            json!([])
        );
        assert_eq!(
            session.core.docx_source.borrow().as_deref(),
            Some(source.as_slice())
        );
        let blocks = |env: &crate::bridge::RenderEnv| {
            session.core.engine.lower_story_json("body", env).unwrap()
        };
        let tokens_env = crate::bridge::RenderEnv {
            media_tokens: true,
            ..Default::default()
        };
        assert!(blocks(&tokens_env).contains(r#""media:0""#));
        assert!(!blocks(&tokens_env).contains("data:"));
        assert!(
            blocks(&crate::bridge::RenderEnv::default()).contains("data:image/png;base64,AQIDBA==")
        );
        assert_eq!(session.media_bytes("media:0"), Some(image_bytes.clone()));
        assert_eq!(session.media_type("media:0").as_deref(), Some("image/png"));
        assert_eq!(
            session.media_data_url("media:0").as_deref(),
            Some("data:image/png;base64,AQIDBA==")
        );
        assert_eq!(session.media_bytes("media:1"), None);
        assert_eq!(
            session.media_data_url("data:image/png;base64,AQIDBA=="),
            None
        );
        let data_url = b"data:image/png;base64,AQIDBA==";
        let holds = |state: &[u8]| {
            state
                .windows(data_url.len())
                .any(|window| window == data_url)
        };
        assert!(holds(
            &session.core.engine.doc().encode_state_as_update_v1()
        ));
        let tokens = EditSession::new(8.0).unwrap();
        tokens.set_media_tokens(true);
        tokens.core.open_docx(&source, true, None, None).unwrap();
        assert!(!holds(
            &tokens.core.engine.doc().encode_state_as_update_v1()
        ));
        assert!(
            tokens
                .core
                .engine
                .lower_story_json("body", &crate::bridge::RenderEnv::default())
                .unwrap()
                .contains(r#""media:0""#)
        );
        assert_eq!(tokens.media_sources_json(), "");
        assert_ne!(session.media_sources_json(), "");
        let materialized: docx_parse::S9WireEnvelope =
            serde_json::from_str(&session.materialize_docx().unwrap().unwrap()).unwrap();
        assert_eq!(materialized, expected);

        let request = serde_json::from_value(json!({
            "determinism": {
                "seed": "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef",
                "now": "2030-01-02T03:04:05.006Z"
            },
            "document": {"content": materialized.document.package.document.content},
            "relationshipEntries": materialized.document.package.relationship_entries,
            "options": {"updateModifiedDate": false}
        }))
        .unwrap();
        let saved = docx_parse::serializer::write_docx_s13(
            request,
            session.core.docx_source.borrow().as_deref().unwrap(),
        )
        .unwrap();
        let saved_parts = ooxml_opc::unzip_parts(&saved).unwrap();
        assert_eq!(
            saved_parts
                .iter()
                .find(|(path, _)| path == "word/media/original.png")
                .unwrap()
                .1,
            image_bytes
        );
        let reopened = crate::seed::parse_docx_for_edit(&saved).unwrap();
        assert_eq!(
            reopened.document.package.media_entries,
            expected.document.package.media_entries
        );
    }

    fn batch_docx() -> Vec<u8> {
        ooxml_opc::rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec()),
            ("word/_rels/document.xml.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#.to_vec()),
            ("word/styles.xml".to_owned(), br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"#.to_vec()),
            ("word/document.xml".to_owned(), br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Alpha</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>beta</w:t></w:r></w:p><w:p w14:paraId="00000002"><w:r><w:t>Gamma</w:t></w:r></w:p></w:body></w:document>"#.to_vec()),
        ])
        .unwrap()
    }

    fn envelope(json: &str) -> Value {
        serde_json::from_str(json).unwrap()
    }

    fn identities_cache_buffer(session: &EditSession) -> *const u8 {
        session
            .core
            .paragraph_identities_cache
            .borrow()
            .as_ref()
            .unwrap()
            .json
            .as_ptr()
    }

    fn checked_paragraph_identities(session: &EditSession) -> String {
        let json = session.paragraph_identities().unwrap();
        assert_eq!(json, session.core.paragraph_identities_uncached().unwrap());
        let buffer = identities_cache_buffer(session);
        assert_eq!(json, session.paragraph_identities().unwrap());
        assert_eq!(buffer, identities_cache_buffer(session));
        let cache = session.core.paragraph_identities_cache.borrow();
        let cached = cache.as_ref().unwrap();
        assert_eq!(cached.version, session.core.engine.doc().version());
        assert_eq!(cached.json, json);
        json
    }

    fn changed_paragraph_identities(session: &EditSession, change: impl FnOnce()) -> String {
        let buffer = identities_cache_buffer(session);
        change();
        let json = checked_paragraph_identities(session);
        assert_ne!(buffer, identities_cache_buffer(session));
        json
    }

    #[test]
    fn paragraph_identities_cache_matches_fresh_json_after_mutations() {
        let mut session = EditSession::new(71.0).unwrap();
        session.core.undo = UndoSession::with_clock(Arc::new(|| 0));
        checked_paragraph_identities(&session);
        changed_paragraph_identities(&session, || {
            session
                .core
                .engine
                .doc()
                .create_story_with_paragraph_id("body", "p1", "abcd", "Normal", "left")
                .unwrap();
        });
        changed_paragraph_identities(&session, || {
            session.begin_opening(Some("first".to_owned()));
        });
        session.track_undo();
        let typed = changed_paragraph_identities(&session, || {
            session
                .insert_text("body", "p1", 0, "x", None, None)
                .unwrap();
        });
        session.add_undo_boundary();
        let split = changed_paragraph_identities(&session, || {
            session
                .split_paragraph("body", "p1", 2, None, None)
                .unwrap();
        });
        assert_ne!(split, typed);
        assert_eq!(
            changed_paragraph_identities(&session, || assert!(session.undo())),
            typed
        );
        assert_eq!(
            changed_paragraph_identities(&session, || assert!(session.redo())),
            split
        );
        changed_paragraph_identities(&session, || {
            session.merge_paragraphs("body", "p1", None, None).unwrap();
        });

        let remote = EditSession::new(72.0).unwrap();
        remote.load(&session.encode_state()).unwrap();
        remote.split_paragraph("body", "p1", 1, None, None).unwrap();
        let update = remote.encode_diff(&session.encode_state_vector()).unwrap();
        changed_paragraph_identities(&session, || session.apply_update(&update).unwrap());
        changed_paragraph_identities(&session, || {
            let receipt = envelope(&session.persist_paragraph_ids().unwrap());
            assert!(!receipt["assignments"].as_array().unwrap().is_empty());
        });
        let saved: Vec<_> = session
            .core
            .engine
            .doc()
            .paragraph_identities()
            .paragraphs
            .into_iter()
            .filter_map(|paragraph| match paragraph.paragraph {
                ParagraphRef::Session { para_id, .. } => {
                    paragraph.ooxml_para_id.map(|id| (para_id, id))
                }
                ParagraphRef::Source(_) => None,
            })
            .collect();
        let saved = serde_json::to_string(&saved).unwrap();
        changed_paragraph_identities(&session, || {
            assert_eq!(session.record_saved_paragraph_ids(&saved).unwrap(), "[]");
        });
        let buffer = identities_cache_buffer(&session);
        session.record_saved_paragraph_ids(&saved).unwrap();
        checked_paragraph_identities(&session);
        assert_eq!(buffer, identities_cache_buffer(&session));
        let before = checked_paragraph_identities(&session);
        assert_eq!(
            changed_paragraph_identities(&session, || {
                session.core.engine.doc().rotate_version(js_entropy());
            }),
            before
        );
        let state = session.encode_state();
        assert_eq!(
            changed_paragraph_identities(&session, || session.load(&state).unwrap()),
            before
        );
    }

    #[test]
    fn paragraph_identities_cache_tracks_source_replacement() {
        let session = EditSession::new(73.0).unwrap();
        let source = batch_docx();
        checked_paragraph_identities(&session);
        let opened = changed_paragraph_identities(&session, || {
            session.core.open_docx(&source, true, None, None).unwrap();
        });
        let version = session.version();
        assert_eq!(
            changed_paragraph_identities(&session, || {
                session.core.engine.doc().retain_source_docx(source.clone());
            }),
            opened
        );
        assert_eq!(session.version(), version);

        let mut parts = ooxml_opc::unzip_parts(&source).unwrap();
        let (_, document) = parts
            .iter_mut()
            .find(|(name, _)| name == "word/document.xml")
            .unwrap();
        *document = String::from_utf8(document.clone())
            .unwrap()
            .replace("00000002", "ABCDEF01")
            .into_bytes();
        let replacement = ooxml_opc::rezip_parts(&parts).unwrap();
        let replaced = changed_paragraph_identities(&session, || {
            session
                .core
                .engine
                .doc()
                .retain_source_docx(replacement.clone());
        });
        assert_ne!(opened, replaced);
        assert_eq!(session.version(), version);
        changed_paragraph_identities(&session, || {
            session
                .core
                .engine
                .doc()
                .retain_source_docx(&b"invalid docx"[..]);
        });
        assert_eq!(
            envelope(&session.paragraph_identities().unwrap())["packageSha256"],
            Value::Null
        );
    }

    #[test]
    fn paragraph_identities_cache_survives_reads_and_layout_slices() {
        let session = EditSession::new(74.0).unwrap();
        session
            .core
            .open_docx(&batch_docx(), true, None, None)
            .unwrap();
        session
            .insert_text(
                "body",
                "00000001",
                0,
                "x",
                Some("Ada".to_owned()),
                Some("2026-09-24T00:00:00Z".to_owned()),
            )
            .unwrap();
        let json = checked_paragraph_identities(&session);
        let buffer = identities_cache_buffer(&session);
        let version = session.version();
        let state = session.encode_state();
        let check = || {
            assert_eq!(checked_paragraph_identities(&session), json);
            assert_eq!(identities_cache_buffer(&session), buffer);
            assert_eq!(session.version(), version);
        };
        session.paragraphs("body").unwrap();
        check();
        session
            .read_paragraphs_json(r#"{"view":"accepted"}"#)
            .unwrap();
        check();
        for paragraph in envelope(&json)["paragraphs"].as_array().unwrap() {
            for anchor in ["session", "source", "persisted"] {
                assert_eq!(
                    envelope(
                        &session
                            .resolve_paragraph_anchor(&paragraph[anchor].to_string())
                            .unwrap()
                    )["status"],
                    "found"
                );
                check();
            }
        }
        assert!(
            !envelope(&session.list_revisions().unwrap())
                .as_array()
                .unwrap()
                .is_empty()
        );
        check();
        session.core.engine.doc().forget_seen();
        let txn = session.core.engine.doc().yrs_doc().transact();
        session.core.engine.doc().with_seen(&txn, |_| ());
        drop(txn);
        check();

        let font = session
            .register_measure_font(include_bytes!(
                "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
            ))
            .unwrap();
        let mut request = json!({
            "bodyStory": "body",
            "regions": {"sections": [{"properties": {}}]},
            "renderEnv": {},
        });
        let requirements = envelope(
            &session
                .layout_font_requirements_json(&request.to_string())
                .unwrap(),
        );
        let chains: serde_json::Map<String, Value> = requirements
            .as_array()
            .unwrap()
            .iter()
            .map(|requirement| {
                (
                    requirement["key"].as_str().unwrap().to_owned(),
                    json!([font]),
                )
            })
            .collect();
        request["measurement"] = json!({
            "fontChains": chains,
            "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
        });
        check();
        let request = request.to_string();
        let mut progress = envelope(&session.begin_region_layout(&request).unwrap());
        check();
        while progress["layoutJson"].is_null() {
            progress = envelope(&session.resume_region_layout(1).unwrap());
            check();
        }
        session
            .layout_document_with_regions_retained_json(&request)
            .unwrap();
        check();
        session.retained_kernel_inputs_json().unwrap();
        check();
        assert_eq!(session.encode_state(), state);
    }

    fn replace_request(version: &str, extra: Value) -> String {
        let mut request = json!({
            "expectVersion": version,
            "steps": [{
                "op": "replaceText",
                "target": {"kind": "search", "text": "beta",
                    "within": {"kind": "paragraph", "story": "body", "paraId": "00000001"},
                    "view": "accepted"},
                "text": "BETA"
            }]
        });
        if let (Some(request), Value::Object(extra)) = (request.as_object_mut(), extra) {
            request.extend(extra);
        }
        request.to_string()
    }

    #[test]
    fn batch_envelopes_default_and_refuse_as_data() {
        let session = EditSession::new(71.0).unwrap();
        session
            .core
            .open_docx(&batch_docx(), true, None, None)
            .unwrap();
        let version = session.version();
        let read = envelope(
            &session
                .read_paragraphs_json(r#"{"view":"accepted"}"#)
                .unwrap(),
        );
        assert_eq!(read["ok"], true);
        assert_eq!(read["version"], version.as_str());
        assert_eq!(read["paragraphs"][0]["text"], "Alpha\u{FFFC}beta");
        assert_eq!(read["paragraphs"][0]["atoms"][0]["kind"], "lineBreak");

        let validated = envelope(
            &session
                .validate_edits_json(&replace_request(&version, json!({})))
                .unwrap(),
        );
        assert_eq!(validated["ok"], true);
        assert_eq!(validated["wouldApply"], true);
        assert_eq!(session.version(), version);

        let applied = envelope(
            &session
                .apply_edits_json(&replace_request(&version, json!({"history": "none"})))
                .unwrap(),
        );
        assert_eq!(applied["ok"], true);
        assert_eq!(applied["applied"], true);
        assert_eq!(applied["source"], "host");
        assert_eq!(applied["baseVersion"], version.as_str());
        assert_eq!(applied["version"], session.version().as_str());
        assert_eq!(applied["changedStories"], json!(["body"]));
        assert_eq!(applied["receipts"][0]["range"]["start"]["offset"], 6);

        let stale = envelope(
            &session
                .apply_edits_json(&replace_request(&version, json!({})))
                .unwrap(),
        );
        assert_eq!(stale["ok"], false);
        assert_eq!(stale["failure"]["code"], "stale-version");
        assert_eq!(stale["version"], session.version().as_str());
        assert_eq!(
            session.core.engine.doc().paragraphs("body").unwrap()[0].text,
            "AlphaBETA"
        );
    }

    #[test]
    fn export_envelopes_carry_versions_and_refusals_as_data() {
        let session = EditSession::new(73.0).unwrap();
        session
            .core
            .open_docx(&batch_docx(), true, None, None)
            .unwrap();
        let version = session.version();
        let read = envelope(
            &session
                .export_structured_json(r#"{"revisionView":"accepted"}"#)
                .unwrap(),
        );
        assert_eq!(read["ok"], true);
        assert_eq!(read["version"], version.as_str());
        assert_eq!(read["content"]["schemaVersion"], 1);
        assert_eq!(read["content"]["anchorScope"], "session");
        assert_eq!(
            read["content"]["stories"][0]["blocks"][0]["paragraph"]["inlines"][1]["kind"],
            "break"
        );
        let markdown = envelope(
            &session
                .export_markdown_json(r#"{"revisionView":"markup","includeFormatting":false}"#)
                .unwrap(),
        );
        assert_eq!(markdown["ok"], true);
        assert!(
            markdown["content"]["markdown"]
                .as_str()
                .unwrap()
                .contains("Alpha\\\nbeta")
        );
        let refused = envelope(
            &session
                .export_structured_json(r#"{"revisionView":"accepted","maxBytes":8}"#)
                .unwrap(),
        );
        assert_eq!(refused["ok"], false);
        assert_eq!(refused["version"], version.as_str());
        assert_eq!(refused["failure"]["code"], "invalid-options");
        assert_eq!(refused["failure"]["target"], Value::Null);
        assert!(
            refused["failure"]
                .as_object()
                .unwrap()
                .contains_key("target")
        );
        assert_eq!(session.version(), version);
        let empty = EditSession::new(74.0).unwrap();
        let unsupported = envelope(
            &empty
                .export_structured_json(r#"{"revisionView":"accepted"}"#)
                .unwrap(),
        );
        assert_eq!(unsupported["failure"]["code"], "unsupported");
        assert_eq!(unsupported["failure"]["target"], Value::Null);
    }

    #[test]
    fn paged_export_envelopes_carry_versions_maps_and_refusals() {
        let font: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        let session = EditSession::new(75.0).unwrap();
        session
            .core
            .open_docx(&batch_docx(), true, None, None)
            .unwrap();
        let options = r#"{"revisionView":"markup"}"#;
        let unavailable = envelope(
            &session
                .export_structured_with_pages_json(options, None)
                .unwrap(),
        );
        assert_eq!(unavailable["ok"], false);
        assert_eq!(unavailable["version"], session.version().as_str());
        assert_eq!(unavailable["failure"]["code"], "layout-unavailable");
        assert_eq!(unavailable["failure"]["target"], Value::Null);
        let mut request = json!({
            "bodyStory": "body",
            "regions": {"sections": [{"properties": {}}]},
            "renderEnv": {},
        });
        let requirements: Vec<Value> = serde_json::from_str(
            &session
                .layout_font_requirements_json(&request.to_string())
                .unwrap(),
        )
        .unwrap();
        let chains = |id: u32| -> serde_json::Map<String, Value> {
            requirements
                .iter()
                .map(|requirement| (requirement["key"].as_str().unwrap().to_owned(), json!([id])))
                .collect()
        };
        let defaults = json!({"fontSize": 11, "fontFamily": "Calibri"});
        let snapshot_request = {
            let mut private = request.clone();
            private["measurement"] = json!({"fontChains": chains(0), "defaults": defaults});
            private.to_string()
        };
        let id = session.register_measure_font(font).unwrap();
        request["measurement"] = json!({"fontChains": chains(id), "defaults": defaults});
        let request = request.to_string();
        session
            .layout_document_with_regions_retained_json(&request)
            .unwrap();
        let version = session.version();
        let read = envelope(
            &session
                .export_structured_with_pages_json(options, Some(request.clone()))
                .unwrap(),
        );
        assert_eq!(read["ok"], true);
        assert_eq!(read["version"], version.as_str());
        assert_eq!(
            read["content"]["layout"]["documentVersion"],
            version.as_str()
        );
        assert_eq!(read["content"]["layout"]["layoutRevisionView"], "markup");
        assert_eq!(read["content"]["layout"]["pages"][0]["displayedLabel"], "1");
        assert!(
            read["content"]["layout"]["provenance"]
                .get("frameEpoch")
                .is_none()
        );
        assert_eq!(
            read["content"]["layout"]["fragments"][0]["geometry"],
            Value::Null
        );
        assert_eq!(read["content"]["structured"]["anchorScope"], "session");
        let private = EditSession::new(76.0).unwrap();
        private
            .core
            .open_docx(&batch_docx(), true, None, None)
            .unwrap();
        let snapshot = envelope(
            &private
                .export_snapshot_with_private_fonts_json(
                    font,
                    &[font.len() as u32],
                    &snapshot_request,
                    options,
                )
                .unwrap(),
        );
        assert_eq!(snapshot["ok"], true);
        assert!(snapshot.get("version").is_none());
        assert!(
            snapshot["content"]["layout"]
                .get("documentVersion")
                .is_none()
        );
        assert_eq!(snapshot["content"]["structured"]["anchorScope"], "snapshot");
        assert_eq!(
            session
                .export_structured_with_pages_json(options, Some(request.clone()))
                .map(|reply| envelope(&reply)["ok"].clone())
                .unwrap(),
            true,
            "the private fonts never replaced the session's"
        );
        let marked = envelope(
            &render_docx_markdown_with_pages_json(
                &snapshot["content"].to_string(),
                r#"{"pageMarkers":true}"#,
            )
            .unwrap(),
        );
        assert!(
            marked["content"]["markdown"]
                .as_str()
                .unwrap()
                .contains("<!-- docx-export:0 --><!-- docx-pages: 0=1 -->")
        );
        let limited = envelope(
            &session
                .export_structured_with_pages_json(
                    r#"{"revisionView":"markup","maxLayoutBytes":8}"#,
                    None,
                )
                .unwrap(),
        );
        assert_eq!(limited["failure"]["code"], "invalid-options");
        assert_eq!(session.version(), version);
    }

    #[test]
    fn every_font_entry_point_reads_the_session_store() {
        const LIBERATION: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        const CARLITO: &[u8] = include_bytes!("../../docx-raster/tests/assets/Carlito-Regular.ttf");
        let parse = |json: String| serde_json::from_str::<Value>(&json).unwrap();
        let open = |client_id: f64| {
            let session = EditSession::new(client_id).unwrap();
            session
                .core
                .open_docx(&batch_docx(), true, None, None)
                .unwrap();
            session
        };
        let request = |session: &EditSession| {
            let mut request = json!({
                "bodyStory": "body",
                "regions": {"sections": [{"properties": {}}]},
                "renderEnv": {},
            });
            let requirements = parse(
                session
                    .layout_font_requirements_json(&request.to_string())
                    .unwrap(),
            );
            let chains: serde_json::Map<String, Value> = requirements
                .as_array()
                .unwrap()
                .iter()
                .map(|requirement| {
                    (
                        requirement["key"].as_str().unwrap().to_owned(),
                        json!([0, 1]),
                    )
                })
                .collect();
            request["measurement"] = json!({
                "fontChains": chains,
                "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
                "authoritativeShaping": true,
            });
            (request.to_string(), chains)
        };
        let drive = |session: &EditSession, fonts: [&[u8]; 2], between: &dyn Fn()| {
            let mut out = Vec::new();
            let mut record = |value: Value| {
                out.push(value);
                between();
            };
            session.clear_measure_fonts();
            record(json!(session.register_measure_font(fonts[0]).unwrap()));
            record(json!(session.register_measure_font(fonts[1]).unwrap()));
            let substitute = session
                .register_substitute_measure_font(0, "MS Mincho")
                .unwrap();
            assert_eq!(substitute, 2);
            record(json!(substitute));
            let (request, chains) = request(session);
            let blocks = parse(session.yrs_blocks_for_story("body", "{}").unwrap());
            record(blocks.clone());
            let measure = json!({
                "block": blocks[0],
                "maxWidth": 200,
                "fontChains": chains,
                "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
                "authoritativeShaping": true,
            });
            record(parse(
                session
                    .measure_paragraph_json(&measure.to_string())
                    .unwrap(),
            ));
            record(parse(
                session.layout_document_with_regions_json(&request).unwrap(),
            ));
            record(parse(
                session
                    .layout_document_with_regions_prefix_retained_json(&request, 1)
                    .unwrap(),
            ));
            let mut progress = parse(session.begin_region_layout(&request).unwrap());
            while progress.get("layoutJson").is_none() {
                between();
                progress = parse(session.resume_region_layout(1).unwrap());
            }
            record(progress);
            let layout = parse(
                session
                    .layout_document_with_regions_retained_json(&request)
                    .unwrap(),
            );
            record(layout.clone());
            let mut kernel = parse(session.retained_kernel_inputs_json().unwrap());
            record(kernel.clone());
            session.reset_frame_base();
            record(json!(session.build_display_list_frame("{}", 0.0).unwrap()));
            session
                .set_selection("body", "00000002", 2, "00000002", 2)
                .unwrap();
            record(parse(session.resident_caret_snapshot_json().unwrap()));
            record(json!(session.apply_input("x", 1.0).unwrap()));
            record(json!(session.apply_delete("backward", 2.0, 1).unwrap()));
            record(parse(
                session
                    .display_hit_test_regions_json(0, 120.0, 100.0)
                    .unwrap(),
            ));
            record(parse(
                session
                    .display_vertical_move_json(3.0, "down", 120.0)
                    .unwrap(),
            ));
            record(parse(session.display_range_rects_json(1.0, 9.0).unwrap()));
            record(parse(
                session
                    .display_range_rects_region_json("body", "", 1.0, 9.0)
                    .unwrap(),
            ));
            record(parse(session.outline_glyph_json(0, 36).unwrap()));
            record(parse(session.outline_glyph_json(1, 36).unwrap()));
            let pages = parse(
                session
                    .export_structured_with_pages_json(
                        r#"{"revisionView":"markup"}"#,
                        Some(request.clone()),
                    )
                    .unwrap(),
            );
            record(pages["content"]["layout"]["fragments"].clone());
            kernel["layout"] = layout["layout"].clone();
            record(parse(
                session
                    .build_display_list_json(&kernel.to_string())
                    .unwrap(),
            ));
            kernel.as_object_mut().unwrap().remove("layout");
            record(parse(
                session.layout_document_json(&kernel.to_string()).unwrap(),
            ));
            let private = parse(
                session
                    .export_snapshot_with_private_fonts_json(
                        fonts[1],
                        &[fonts[1].len() as u32],
                        &request.replace("[0,1]", "[0]"),
                        r#"{"revisionView":"markup"}"#,
                    )
                    .unwrap(),
            );
            record(private["content"]["layout"]["fragments"].clone());
            record(parse(session.outline_glyph_json(0, 36).unwrap()));
            session.set_display_window(0, 0);
            session
                .layout_document_with_regions_retained_json(&request)
                .unwrap();
            record(json!(session.build_display_list_frame("{}", 0.0).unwrap()));
            record(json!(
                session.build_display_pages_frame(vec![0], 0.0).unwrap()
            ));
            out
        };
        let quiet = || {};
        let a_alone = drive(&open(81.0), [LIBERATION, CARLITO], &quiet);
        let b_alone = drive(&open(82.0), [CARLITO, LIBERATION], &quiet);
        assert_ne!(a_alone, b_alone);

        let b = open(84.0);
        let (b_request, _) = request(&b);
        let b_round = || {
            b.clear_measure_fonts();
            b.register_measure_font(CARLITO).unwrap();
            b.register_measure_font(LIBERATION).unwrap();
            b.layout_document_with_regions_retained_json(&b_request)
                .unwrap();
            b.build_display_list_frame("{}", 0.0).unwrap();
            b.outline_glyph_json(1, 36).unwrap();
        };
        b_round();
        assert_eq!(drive(&open(83.0), [LIBERATION, CARLITO], &b_round), a_alone);
    }

    #[test]
    fn snapshot_exports_have_no_version_and_render_from_content() {
        let options = r#"{"revisionView":"accepted"}"#;
        let exported = envelope(&export_docx_structured_json(&batch_docx(), options).unwrap());
        assert_eq!(exported["ok"], true);
        assert!(exported.get("version").is_none());
        assert_eq!(exported["content"]["anchorScope"], "snapshot");
        let rendered =
            envelope(&render_docx_markdown_json(&exported["content"].to_string(), "{}").unwrap());
        let markdown = envelope(&export_docx_markdown_json(&batch_docx(), options).unwrap());
        assert_eq!(rendered, markdown);
        let refused = envelope(
            &render_docx_markdown_json(&exported["content"].to_string(), r#"{"maxBytes":1}"#)
                .unwrap(),
        );
        assert_eq!(refused["failure"]["code"], "invalid-options");
    }

    #[test]
    fn untracked_batches_keep_history_and_source_is_echoed() {
        let session = EditSession::new(72.0).unwrap();
        session
            .core
            .open_docx(&batch_docx(), true, None, None)
            .unwrap();
        let request = replace_request(
            &session.version(),
            json!({"source": "agent", "history": "none"}),
        );
        let applied = envelope(&session.apply_edits_json(&request).unwrap());
        assert_eq!(applied["source"], "agent");
        assert!(!session.can_undo());
    }

    /// A package whose paragraph text id is invalid, so parsing generates
    /// one from the package digest.
    fn fixture_packages() -> Vec<(std::path::PathBuf, Vec<u8>)> {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"));
        let mut pending = vec![root.join("tests/fixtures")];
        let mut found = Vec::new();
        while let Some(dir) = pending.pop() {
            for entry in std::fs::read_dir(dir).unwrap() {
                let path = entry.unwrap().path();
                if path.is_dir() {
                    pending.push(path);
                } else if path
                    .extension()
                    .is_some_and(|extension| extension == "docx")
                {
                    let bytes = std::fs::read(&path).unwrap();
                    found.push((path, bytes));
                }
            }
        }
        found.sort();
        found
    }

    #[test]
    fn an_unchanged_session_plans_its_save_from_what_a_reseed_would_hold() {
        let mut checked = 0;
        for (path, bytes) in fixture_packages() {
            let label = path.display().to_string();
            let session = EditSession::new(81.0).unwrap();
            if session
                .core
                .open_docx_retaining(&bytes, true, None, None)
                .is_err()
            {
                continue;
            }
            let doc = session.core.engine.doc();
            assert!(doc.unchanged_since_seeded(), "{label}");
            let reused = doc.paragraph_save_plan();
            crate::identity::assert_seed_states_match_a_reseed(doc, &label);
            let reseeding = EditSession::new(82.0).unwrap();
            reseeding
                .core
                .open_docx_retaining(&bytes, true, None, None)
                .unwrap();
            reseeding.core.engine.doc().forget_seeding();
            assert_eq!(
                reseeding.core.engine.doc().paragraph_save_plan(),
                reused,
                "{label}"
            );
            checked += 1;
        }
        assert!(checked >= 50, "{checked} fixtures opened");
    }

    #[test]
    fn an_open_into_a_session_holding_stories_plans_from_a_reseed() {
        let bytes = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("tests/fixtures/footnote-anchor.docx"),
        )
        .unwrap();
        let opened = |client: f64| {
            let session = EditSession::new(client).unwrap();
            session
                .create_story("fn:99", "Existing note", "Normal", "left")
                .unwrap();
            session
                .core
                .open_docx_retaining(&bytes, true, None, None)
                .unwrap();
            session
        };
        let session = opened(86.0);
        assert!(!session.core.engine.doc().unchanged_since_seeded());
        let plan = session.core.engine.doc().paragraph_save_plan();
        let reseeding = opened(87.0);
        reseeding.core.engine.doc().forget_seeding();
        assert_eq!(reseeding.core.engine.doc().paragraph_save_plan(), plan);
    }

    #[test]
    fn seeding_comment_metadata_keeps_a_session_planning_from_its_own_stories() {
        let bytes = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("tests/fixtures/structured-export/principal.docx"),
        )
        .unwrap();
        let opened = |client: f64| {
            let session = EditSession::new(client).unwrap();
            session
                .core
                .open_docx_retaining(&bytes, true, None, None)
                .unwrap();
            let comments: Vec<Value> =
                serde_json::from_str(&session.list_comments().unwrap()).unwrap();
            assert!(!comments.is_empty());
            for comment in &comments {
                let id = match &comment["id"] {
                    Value::String(id) => id.clone(),
                    id => id.to_string(),
                };
                session
                    .seed_comment_metadata_json(
                        &id,
                        r#"{"author":"A","date":"2026-01-01T00:00:00Z","done":true,"parentId":null,"body":null}"#,
                    )
                    .unwrap();
            }
            session
                .seed_comment_metadata_json("reply-1", r#"{"author":"B","parentId":"1"}"#)
                .unwrap();
            session
        };
        let session = opened(88.0);
        assert!(session.core.engine.doc().unchanged_since_seeded());
        let plan = session.core.engine.doc().paragraph_save_plan();
        crate::identity::assert_seed_states_match_a_reseed(session.core.engine.doc(), "principal");
        let reseeding = opened(89.0);
        reseeding.core.engine.doc().forget_seeding();
        assert_eq!(reseeding.core.engine.doc().paragraph_save_plan(), plan);
    }

    #[test]
    fn an_open_under_another_digest_plans_from_a_reseed() {
        let bytes = unidentified_docx("Alpha");
        let digest = "0".repeat(64);
        let opened = |client: f64| {
            let session = EditSession::new(client).unwrap();
            session
                .core
                .open_docx_retaining(&bytes, true, Some("fixed"), Some(&digest))
                .unwrap();
            session
        };
        let session = opened(91.0);
        let plan = session.core.engine.doc().paragraph_save_plan();
        let reseeding = opened(92.0);
        reseeding.core.engine.doc().forget_seeding();
        assert_eq!(reseeding.core.engine.doc().paragraph_save_plan(), plan);
    }

    #[test]
    fn a_session_edited_since_seeding_plans_from_a_reseed() {
        let session = EditSession::new(83.0).unwrap();
        session
            .core
            .open_docx_retaining(&unidentified_docx("Alpha"), true, None, None)
            .unwrap();
        assert!(session.core.engine.doc().unchanged_since_seeded());
        let para_id = session.core.engine.doc().paragraphs("body").unwrap()[0]
            .para_id
            .clone();
        session
            .insert_text("body", &para_id, 1, "!", None, None)
            .unwrap();
        assert!(!session.core.engine.doc().unchanged_since_seeded());
        session.core.engine.doc().paragraph_save_plan();
        crate::identity::assert_seed_states_match_a_reseed(session.core.engine.doc(), "edited");

        let restored = EditSession::new(84.0).unwrap();
        restored
            .core
            .open_docx_retaining(&unidentified_docx("Alpha"), true, None, None)
            .unwrap();
        let para_id = restored.core.engine.doc().paragraphs("body").unwrap()[0]
            .para_id
            .clone();
        restored
            .insert_text("body", &para_id, 1, "!", None, None)
            .unwrap();
        restored
            .delete_range("body", &para_id, 1, &para_id, 2, None, None)
            .unwrap();
        assert_eq!(
            restored.core.engine.doc().paragraphs("body").unwrap()[0].text,
            "Alpha"
        );
        assert!(!restored.core.engine.doc().unchanged_since_seeded());

        let remote = EditSession::new(85.0).unwrap();
        remote
            .core
            .open_docx_retaining(&unidentified_docx("Alpha"), true, None, None)
            .unwrap();
        remote.apply_update(&session.encode_state()).unwrap();
        assert!(!remote.core.engine.doc().unchanged_since_seeded());
        remote.core.engine.doc().paragraph_save_plan();
        crate::identity::assert_seed_states_match_a_reseed(remote.core.engine.doc(), "remote");
    }

    fn unidentified_docx(text: &str) -> Vec<u8> {
        let document = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:textId="invalid"><w:r><w:t>{text}</w:t></w:r></w:p></w:body></w:document>"#
        );
        ooxml_opc::rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec()),
            ("word/document.xml".to_owned(), document.into_bytes()),
        ])
        .unwrap()
    }

    fn script_fonts_docx(body: &str, defaults: &str) -> Vec<u8> {
        const MAIN: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape""#;
        let document = format!(r#"<w:document {MAIN}><w:body>{body}</w:body></w:document>"#);
        let styles = format!(
            r#"<w:styles {MAIN}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="SimSun" w:cs="Times New Roman"/>{defaults}</w:rPr></w:rPrDefault></w:docDefaults></w:styles>"#
        );
        let fonts = format!(
            r#"<w:fonts {MAIN}><w:font w:name="Calibri"><w:charset w:val="00"/></w:font><w:font w:name="SimSun"><w:altName w:val="宋体"/><w:charset w:val="86"/></w:font><w:font w:name="Batang"><w:charset w:val="81"/></w:font></w:fonts>"#
        );
        ooxml_opc::rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec()),
            ("word/_rels/document.xml.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/></Relationships>"#.to_vec()),
            ("word/document.xml".to_owned(), document.into_bytes()),
            ("word/styles.xml".to_owned(), styles.into_bytes()),
            ("word/fontTable.xml".to_owned(), fonts.into_bytes()),
        ])
        .unwrap()
    }

    fn unused_script_fonts(body: &str, defaults: &str, seed_stories: bool) -> Value {
        let session = EditSession::new(78.0).unwrap();
        let bytes = script_fonts_docx(body, defaults);
        let host: Value = serde_json::from_str(
            &session
                .core
                .open_docx(&bytes, seed_stories, None, None)
                .unwrap(),
        )
        .unwrap();
        host["unusedScriptFonts"].clone()
    }

    #[test]
    fn an_opened_document_names_the_script_fonts_its_text_does_not_use() {
        let run = |text: &str| format!("<w:r><w:t>{text}</w:t></w:r>");
        let latin = |text: &str| format!("<w:p>{}</w:p><w:p>{text}</w:p>", run("Latin"));
        let all = json!(["Batang", "SimSun", "Times New Roman", "宋体"]);
        assert_eq!(unused_script_fonts(&latin(&run("More")), "", true), all);
        assert_eq!(
            unused_script_fonts(&latin(&run("漢字")), "", true),
            json!(["Times New Roman"])
        );
        assert_eq!(
            unused_script_fonts(&latin(&run("مرحبا")), "", true),
            json!(["Batang", "SimSun", "宋体"])
        );
        assert_eq!(
            unused_script_fonts(&latin(&run("More")), "", false),
            json!([])
        );
        let control = format!(
            "<w:p><w:sdt><w:sdtPr/><w:sdtContent>{}</w:sdtContent></w:sdt></w:p>",
            run("abc")
        );
        assert_eq!(unused_script_fonts(&control, "", true), all);
        assert_eq!(
            unused_script_fonts(&control, "<w:cs/>", true),
            json!(["Batang", "SimSun", "宋体"])
        );
        let text_box = r#"<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Text box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:rFonts w:ascii="SimSun" w:hAnsi="SimSun"/></w:rPr><w:t>Hello</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#;
        assert_eq!(
            unused_script_fonts(&latin(text_box), "", true),
            json!(["Batang", "Times New Roman", "宋体"])
        );
        let vml_text_box = r#"<w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape style="width:72pt;height:36pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>漢字</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>"#;
        assert_eq!(
            unused_script_fonts(&latin(vml_text_box), "", true),
            json!(["Times New Roman"])
        );
    }

    #[test]
    fn a_preview_names_the_script_fonts_its_cut_does_not_use() {
        let paragraph = |text: &str| format!("<w:p><w:r><w:t>{text}</w:t></w:r></w:p>");
        let unused = |body: String| -> Value {
            let session = EditSession::new(79.0).unwrap();
            let host = session
                .core
                .open_preview(&script_fonts_docx(&body, ""), 1)
                .unwrap()
                .unwrap();
            serde_json::from_str::<Value>(&host).unwrap()["unusedScriptFonts"].clone()
        };
        let body = |first: &str, second: &str| format!("{}{}", paragraph(first), paragraph(second));
        let all = json!(["Batang", "SimSun", "Times New Roman", "宋体"]);
        assert_eq!(unused(body("Latin", "More")), all);
        assert_eq!(unused(body("漢字", "Latin")), json!(["Times New Roman"]));
        // The CJK text lies past the preview's cut.
        assert_eq!(unused(body("Latin", "漢字")), all);
    }

    #[test]
    fn a_failed_reopen_keeps_the_open_package_and_its_digest() {
        let opened = unidentified_docx("Alpha");
        let session = EditSession::new(75.0).unwrap();
        session.core.open_docx(&opened, true, None, None).unwrap();
        let materialized = session.materialize_docx().unwrap();
        assert!(
            session
                .core
                .open_docx_retaining(&unidentified_docx("Beta"), true, None, None)
                .is_err()
        );
        assert_eq!(session.materialize_docx().unwrap(), materialized);
        let fresh = EditSession::new(76.0).unwrap();
        fresh.core.open_docx(&opened, true, None, None).unwrap();
        assert_eq!(fresh.materialize_docx().unwrap(), materialized);
    }

    #[test]
    fn an_open_with_the_package_digest_matches_one_that_hashes_the_package() {
        let bytes = unidentified_docx("Alpha");
        let generation = Some("fixed".to_owned());
        let hashed = EditSession::new(77.0).unwrap();
        hashed
            .core
            .open_docx(&bytes, true, generation.clone(), None)
            .unwrap();
        let given = EditSession::new(77.0).unwrap();
        given
            .core
            .open_docx(
                &bytes,
                true,
                generation,
                Some(crate::seed::package_digest(&bytes)),
            )
            .unwrap();
        assert!(hashed.encode_state() == given.encode_state());
        assert_eq!(
            hashed.core.docx_digest.borrow().clone(),
            given.core.docx_digest.borrow().clone()
        );
        assert_eq!(
            hashed.materialize_docx().unwrap(),
            given.materialize_docx().unwrap()
        );
        let other = EditSession::new(77.0).unwrap();
        other
            .core
            .open_docx(&bytes, true, Some("fixed".to_owned()), Some("0".repeat(64)))
            .unwrap();
        assert_ne!(
            hashed.materialize_docx().unwrap(),
            other.materialize_docx().unwrap(),
            "the digest seeds the generated ids"
        );

        let refused = EditSession::new(78.0).unwrap();
        for malformed in ["", "ABC", &"A".repeat(64), &"g".repeat(64)] {
            assert!(
                refused
                    .core
                    .open_docx_retaining(&bytes, true, None, Some(malformed))
                    .is_err()
            );
        }
        assert!(refused.materialize_docx().unwrap().is_none());
    }

    #[test]
    fn a_preview_renders_numpages_empty_however_much_of_it_is_laid_out() {
        let mut body = String::from(
            r#"<w:p><w:r><w:t>Of </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>"#,
        );
        for index in 0..60 {
            body.push_str(&format!(
                "<w:p><w:r><w:t>Paragraph {index}</w:t></w:r></w:p>"
            ));
        }
        let document = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#
        );
        let bytes = ooxml_opc::rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec()),
            ("word/document.xml".to_owned(), document.into_bytes()),
        ])
        .unwrap();
        const LIBERATION: &[u8] =
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        let font_id = 0;
        let request = json!({
            "bodyStory": "body",
            "regions": { "sections": [{ "sectionId": "main", "properties": {
                "pageWidth": 5760, "pageHeight": 2880,
                "marginTop": 360, "marginRight": 360, "marginBottom": 360, "marginLeft": 360
            } }] },
            "measurement": {
                "fontChains": { "liberation sans|0|0": [font_id] },
                "defaults": { "fontSize": 11, "fontFamily": "Liberation Sans" },
                "authoritativeShaping": true
            },
            "renderEnv": {}
        })
        .to_string();
        let numpages = |session: &EditSession| {
            let _fonts = session.core.fonts.enter();
            let layout: Value = serde_json::from_str(
                &session
                    .core
                    .engine
                    .layout_document_with_regions_retained_json(&request)
                    .unwrap(),
            )
            .unwrap();
            session
                .core
                .engine
                .build_display_list_frame("{}", 0)
                .unwrap();
            let text = session
                .core
                .engine
                .with_display_list(|list| {
                    let page = serde_json::to_value(&list.pages[0]).unwrap();
                    page["primitives"]
                        .as_array()
                        .unwrap()
                        .iter()
                        .filter(|primitive| primitive["field"]["category"] == "NUMPAGES")
                        .map(|primitive| primitive["text"].as_str().unwrap_or("").to_owned())
                        .collect::<String>()
                })
                .unwrap();
            (text, layout["layout"]["pages"].as_array().unwrap().len())
        };
        let full = EditSession::new(80.0).unwrap();
        full.register_measure_font(LIBERATION).unwrap();
        full.core.open_docx(&bytes, true, None, None).unwrap();
        let (text, pages) = numpages(&full);
        assert!(pages > 1);
        assert_eq!(text, pages.to_string());
        // The preview's blocks fit its pages whole; its page count is still not the document's.
        let preview = EditSession::new(81.0).unwrap();
        preview.register_measure_font(LIBERATION).unwrap();
        preview.core.open_preview(&bytes, 20).unwrap().unwrap();
        let (text, preview_pages) = numpages(&preview);
        assert!(preview_pages < pages);
        assert_eq!(text, "");

        // An edit's resident pass lays out part of the package too.
        let fonts = preview.core.fonts.enter();
        preview
            .core
            .engine
            .doc()
            .insert_text(
                &EditCtx::local("", ""),
                Position::new("body", 40),
                "x",
                FormatPolicy::Inherit,
            )
            .unwrap();
        let epoch = preview.core.engine.stats().frame_epoch;
        preview.core.engine.apply_and_layout("body", epoch).unwrap();
        assert!(preview.core.engine.stats().incremental_pagination_calls > 0);
        preview
            .core
            .engine
            .build_display_list_frame("{\"changed\":1}", 0)
            .unwrap();
        let text = preview
            .core
            .engine
            .with_display_list(|list| {
                serde_json::to_value(&list.pages[0]).unwrap()["primitives"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .filter(|primitive| primitive["field"]["category"] == "NUMPAGES")
                    .map(|primitive| primitive["text"].as_str().unwrap_or("").to_owned())
                    .collect::<String>()
            })
            .unwrap();
        assert_eq!(text, "");
        drop(fonts);

        // A complete open over the preview counts the whole document again.
        preview.delete_story("body").unwrap();
        preview.core.open_docx(&bytes, true, None, None).unwrap();
        assert_eq!(numpages(&preview), (pages.to_string(), pages));
    }

    #[test]
    fn a_preview_reports_whether_its_cut_holds_the_whole_body() {
        let body = (0..30)
            .map(|index| format!("<w:p><w:r><w:t>Paragraph {index}</w:t></w:r></w:p>"))
            .collect::<String>();
        let document = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#
        );
        let bytes = ooxml_opc::rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec()),
            ("word/document.xml".to_owned(), document.into_bytes()),
        ])
        .unwrap();
        for (blocks, whole) in [(10, false), (30, false), (31, true), (200, true)] {
            let preview = EditSession::new(83.0).unwrap();
            let host: Value =
                serde_json::from_str(&preview.core.open_preview(&bytes, blocks).unwrap().unwrap())
                    .unwrap();
            assert_eq!(host.get("wholeBody").is_some(), whole, "{blocks} blocks");
        }
        let full = EditSession::new(84.0).unwrap();
        let host: Value =
            serde_json::from_str(&full.core.open_docx(&bytes, true, None, None).unwrap()).unwrap();
        assert!(host.get("wholeBody").is_none());
    }

    #[test]
    fn a_weighted_preview_reports_budget_cuts_and_genuine_exhaustion() {
        let p = "<w:p><w:r><w:t>Cell</w:t></w:r></w:p>";
        let table = format!("<w:tbl><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>", p.repeat(10));
        for (count, tail, paragraph_budget, whole) in [
            (60, 170, Some(256), false),
            (60, 170, None, false),
            (60, 0, Some(256), true),
            (20, 0, Some(256), true),
            (32, 0, Some(256), true),
            (40, 0, Some(1000), true),
        ] {
            let bytes = script_fonts_docx(&(table.repeat(count) + &p.repeat(tail)), "");
            let preview = EditSession::new(83.0).unwrap();
            let host: Value = serde_json::from_str(
                &preview
                    .core
                    .open_preview_source(&bytes[..], 200, paragraph_budget)
                    .unwrap()
                    .unwrap(),
            )
            .unwrap();
            assert_eq!(
                host.get("wholeBody").is_some(),
                whole,
                "{count} tables, {tail} paragraphs, budget={paragraph_budget:?}"
            );
        }
    }

    #[test]
    fn a_preview_refuses_a_document_whose_later_body_reaches_its_first_pages() {
        let anchor = r#"<w:p><w:r><w:drawing><wp:anchor simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="457200"/><wp:wrapTopAndBottom/><wp:docPr id="1" name="Float"/></wp:anchor></w:drawing></w:r></w:p>"#;
        let section_break = r#"<w:p><w:pPr><w:sectPr/></w:pPr></w:p>"#;
        let columns = r#"<w:sectPr><w:cols w:num="2"/></w:sectPr>"#;
        for (late, end) in [(anchor, ""), (section_break, columns)] {
            let body = (0..200)
                .map(|index| {
                    let late = if index == 149 { late } else { "" };
                    format!("{late}<w:p><w:r><w:t>Paragraph {index}</w:t></w:r></w:p>")
                })
                .collect::<String>();
            let document = format!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>{body}{end}</w:body></w:document>"#
            );
            let bytes = ooxml_opc::rezip_parts(&[
                ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
                ("_rels/.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec()),
                ("word/document.xml".to_owned(), document.into_bytes()),
            ])
            .unwrap();
            let preview = EditSession::new(82.0).unwrap();
            assert!(preview.core.open_preview(&bytes, 60).unwrap().is_none());
            assert!(preview.story_ids().is_empty());
            preview.core.open_docx(&bytes, true, None, None).unwrap();
        }
    }

    #[test]
    fn a_preview_opens_only_into_an_empty_session() {
        let bytes = batch_docx();
        let opened = EditSession::new(77.0).unwrap();
        opened.core.open_docx(&bytes, false, None, None).unwrap();
        assert!(opened.core.open_preview(&bytes, 1).is_err());
        let seeded = EditSession::new(78.0).unwrap();
        seeded.core.open_docx(&bytes, true, None, None).unwrap();
        assert!(seeded.core.open_preview(&bytes, 1).is_err());
        let preview = EditSession::new(79.0).unwrap();
        preview.core.open_preview(&bytes, 1).unwrap().unwrap();
        assert!(preview.materialize_docx().unwrap().is_none());
    }

    #[test]
    fn a_replica_indexes_its_package_as_it_opens() {
        let bytes = batch_docx();
        let origin = EditSession::new(83.0).unwrap();
        origin.core.open_docx(&bytes, true, None, None).unwrap();
        let replica = EditSession::new(84.0).unwrap();
        replica.core.open_docx(&bytes, false, None, None).unwrap();
        assert!(replica.core.engine.doc().source_indexed());
        replica.load(&origin.encode_state()).unwrap();
        let paragraphs = |session: &EditSession| {
            envelope(
                &session
                    .read_paragraphs_json(r#"{"story":"body","view":"accepted"}"#)
                    .unwrap(),
            )["paragraphs"]
                .clone()
        };
        assert_eq!(paragraphs(&replica), paragraphs(&origin));
        assert!(replica.materialize_docx().unwrap().is_some());
    }

    #[test]
    fn reopening_and_hydrating_invalidate_versions() {
        let bytes = batch_docx();
        let origin = EditSession::new(73.0).unwrap();
        origin.core.open_docx(&bytes, true, None, None).unwrap();
        let joined = EditSession::new(74.0).unwrap();
        let before_open = joined.version();
        joined.core.open_docx(&bytes, false, None, None).unwrap();
        assert_ne!(joined.version(), before_open);
        let before_load = joined.version();
        joined.load(&origin.encode_state()).unwrap();
        assert_ne!(joined.version(), before_load);
        let before_reopen = joined.version();
        joined.core.open_docx(&bytes, false, None, None).unwrap();
        assert_ne!(joined.version(), before_reopen);
        let structural = json!({
            "expectVersion": joined.version(),
            "history": "none",
            "steps": [{
                "op": "insertParagraphs",
                "target": {"story": "body", "paraId": "00000002"},
                "at": "end",
                "paragraphs": [{"text": "Joined"}]
            }]
        });
        let applied = envelope(&joined.apply_edits_json(&structural.to_string()).unwrap());
        assert_eq!(applied["ok"], true);
        assert_eq!(applied["receipts"][0]["newParagraphs"][0]["story"], "body");
    }

    #[test]
    fn compatibility_helpers_resolve_targets_after_atoms() {
        let session = EditSession::new(75.0).unwrap();
        session
            .core
            .open_docx(&batch_docx(), true, None, None)
            .unwrap();
        let target = r#"{"kind":"search","text":"beta","within":{"kind":"paragraph","story":"body","paraId":"00000001"},"view":"accepted"}"#;
        let formatted = envelope(
            &session
                .format_text_target_json(target, r#"{"bold":true}"#)
                .unwrap(),
        );
        assert_eq!(formatted["ok"], true);
        let bold: String = session
            .core
            .engine
            .doc()
            .story_segments("body")
            .unwrap()
            .into_iter()
            .filter_map(|segment| match segment.content {
                SegmentContent::Text(text)
                    if segment.attributes.get("bold") == Some(&Any::Bool(true)) =>
                {
                    Some(text)
                }
                _ => None,
            })
            .collect();
        assert_eq!(bold, "beta");
        let commented = envelope(
            &session
                .comment_text_target_json(
                    target,
                    r#"{"id":"7","author":"Ann","date":"2026-09-24T00:00:00Z","body":"note"}"#,
                )
                .unwrap(),
        );
        assert_eq!(commented["ok"], true);
        let anchor = session
            .core
            .engine
            .doc()
            .resolve_comment("7")
            .unwrap()
            .remove(0);
        assert_eq!((anchor.start, anchor.end), (6, 10));
        let missing = envelope(
            &session
                .format_text_target_json(
                    r#"{"kind":"paragraph","story":"body","paraId":"missing"}"#,
                    "{}",
                )
                .unwrap(),
        );
        assert_eq!(missing["failure"]["code"], "missing-target");
        let selection = envelope(
            &session
                .selection_text_json("body", "00000001", 3, "00000002", 2)
                .unwrap(),
        );
        assert_eq!(selection["selectedText"], "ha\u{FFFC}beta\nGa");
        assert_eq!(selection["before"], "Alp");
        assert_eq!(selection["after"], "mma");
    }

    #[test]
    fn image_geometry_at_a_position_reaches_one_of_several_images_sharing_an_id() {
        let session = EditSession::new(22.0).unwrap();
        session
            .core
            .engine
            .doc()
            .create_story_with_paragraph_id("body", "p0", "AB", "Normal", "left")
            .unwrap();
        for offset in [2, 1] {
            session
                .insert_image("body", "p0", offset, r#"{"rId":"rIdShared"}"#, None, None)
                .unwrap();
        }
        session
            .set_image_geometry_at("body", "p0", 3, r#"{"alt":"second"}"#)
            .unwrap();
        session
            .set_image_geometry("rIdShared", r#"{"title":"first"}"#)
            .unwrap();

        let images: Vec<_> = session
            .core
            .engine
            .doc()
            .story_segments("body")
            .unwrap()
            .into_iter()
            .filter_map(|segment| match segment.content {
                SegmentContent::OtherEmbed { payload, .. } => Some(payload),
                _ => None,
            })
            .collect();
        assert_eq!(images.len(), 2);
        assert_eq!(images[0].get("alt"), None);
        assert_eq!(images[0].get("title"), Some(&Any::from("first")));
        assert_eq!(images[1].get("alt"), Some(&Any::from("second")));
        assert_eq!(images[1].get("title"), None);
    }

    fn seed_paragraph_after_embeds(doc: &EditingDoc, embeds: &[&str], text: &str) {
        doc.create_story_with_paragraph_id("body", "p0", "Alpha", "Normal", "left")
            .unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        let mut index = 6;
        let mut ops = Vec::new();
        for kind in embeds {
            ops.push(RawOp::InsertEmbed {
                index,
                kind: (*kind).into(),
                payload: Default::default(),
                attrs: Default::default(),
            });
            index += 1;
        }
        ops.push(RawOp::Insert {
            index,
            text: text.into(),
            attrs: Default::default(),
        });
        index += text.encode_utf16().count() as u32;
        ops.push(RawOp::InsertEmbed {
            index,
            kind: "pilcrow".into(),
            payload: vec![
                ("paraId".into(), Any::from("p1")),
                ("pStyle".into(), Any::from("Normal")),
                ("alignment".into(), Any::from("left")),
            ],
            attrs: Default::default(),
        });
        doc.apply_raw_ops("body", ops, &ctx).unwrap();
    }

    fn caret_layout_input() -> String {
        json!({
            "measured": [{
                "block": {
                    "kind": "paragraph",
                    "id": "p1",
                    "paraId": "p1",
                    "runs": [{
                        "kind": "text",
                        "text": "ABCDE",
                        "pmStart": 1,
                        "pmEnd": 6
                    }],
                    "attrs": {},
                    "pmStart": 0,
                    "pmEnd": 7
                },
                "measure": {
                    "kind": "paragraph",
                    "lines": [{
                        "headRun": 0,
                        "headChar": 0,
                        "tailRun": 0,
                        "tailChar": 5,
                        "width": 50,
                        "ascent": 8,
                        "descent": 2,
                        "lineHeight": 20
                    }],
                    "totalHeight": 20
                }
            }],
            "options": {
                "pageSize": { "w": 200, "h": 120 },
                "margins": { "top": 10, "right": 10, "bottom": 10, "left": 10 }
            }
        })
        .to_string()
    }

    #[test]
    fn index_loc_returns_the_node_offset_for_cumulative_block_embeds() {
        let kinds = ["table", "pageBreak", "columnBreak", "blockSdt"];
        for count in 0..=kinds.len() {
            let doc = EditingDoc::new(7 + count as u64);
            seed_paragraph_after_embeds(&doc, &kinds[..count], "ABCDE");
            let index = 6 + count as u32 + 3;
            let loc = index_loc(&doc, "body", index).unwrap();

            assert_eq!(loc.para_id, "p1");
            assert_eq!(loc.offset, count as u32 + 3);
            assert_eq!(loc.node_offset, 3);
        }
    }

    #[test]
    fn resident_backspace_after_a_leading_embed_deletes_the_previous_character() {
        let session = EditSession::new(20.0).unwrap();
        seed_paragraph_after_embeds(session.core.engine.doc(), &["table"], "ABCDE");
        let head = loc_index(session.core.engine.doc(), "body", "p1", 6).unwrap();

        session
            .core
            .delete_resident_input("backward", ("body".into(), "p1".into(), head), true)
            .unwrap();

        let paragraphs = session.core.engine.doc().paragraphs("body").unwrap();
        assert_eq!(paragraphs[1].text, "ABCD");
    }

    fn resident_surrogate_session(client_id: f64) -> EditSession {
        let session = EditSession::new(client_id).unwrap();
        let doc = session.core.engine.doc();
        doc.create_story_with_paragraph_id("body", "p0", "a😀b", "Normal", "left")
            .unwrap();
        doc.split_paragraph(&EditCtx::local("", ""), Position::new("body", 3), None)
            .unwrap();
        let font_id = session
            .register_measure_font(include_bytes!(
                "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
            ))
            .unwrap();
        let request = json!({
            "bodyStory": "body",
            "regions": {"sections": [{
                "sectionId": "main",
                "properties": {
                    "pageWidth": 4320,
                    "pageHeight": 2880,
                    "marginTop": 300,
                    "marginRight": 300,
                    "marginBottom": 300,
                    "marginLeft": 300
                }
            }]},
            "measurement": {
                "fontChains": {"calibri|0|0": [font_id]},
                "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
                "authoritativeShaping": true
            },
            "renderEnv": {}
        });
        session
            .layout_document_with_regions_json(&request.to_string())
            .unwrap();
        session
            .build_display_list_frame(
                &json!({"fontChains": {"calibri|0|0": [font_id]}}).to_string(),
                0.0,
            )
            .unwrap();
        session.set_selection("body", "p0", 2, "p0", 2).unwrap();
        assert_eq!(
            session.core.collapsed_resident_input_selection().unwrap(),
            ("body".into(), "p0".into(), 2)
        );
        session
    }

    #[track_caller]
    fn assert_resident_surrogate_texts(session: &EditSession, expected: &[&str]) {
        let texts = |doc: &EditingDoc| {
            doc.paragraphs("body")
                .unwrap()
                .into_iter()
                .map(|paragraph| paragraph.text)
                .collect::<Vec<_>>()
        };
        let doc = session.core.engine.doc();
        assert_eq!(texts(doc), expected);
        let peer = EditingDoc::new(99);
        peer.apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(texts(&peer), texts(doc));
    }

    #[test]
    fn resident_backspace_inside_a_surrogate_pair_deletes_the_whole_emoji() {
        let session = resident_surrogate_session(23.0);
        session.core.delete_resident_units("backward", 1).unwrap();
        assert_resident_surrogate_texts(&session, &["a", "b"]);
    }

    #[test]
    fn resident_forward_delete_inside_a_surrogate_pair_deletes_the_whole_emoji() {
        let session = resident_surrogate_session(24.0);
        session.core.delete_resident_units("forward", 1).unwrap();
        assert_resident_surrogate_texts(&session, &["a", "b"]);
    }

    #[test]
    fn resident_typing_inside_a_surrogate_pair_lands_before_the_emoji() {
        let session = resident_surrogate_session(25.0);
        session.apply_input("x", 1.0).unwrap();
        assert_resident_surrogate_texts(&session, &["ax😀", "b"]);
        assert_eq!(
            session.core.collapsed_resident_input_selection().unwrap(),
            ("body".into(), "p0".into(), 2)
        );
        session.core.delete_resident_units("backward", 1).unwrap();
        assert_resident_surrogate_texts(&session, &["a😀", "b"]);
    }

    #[test]
    fn resident_caret_snapshot_returns_null_for_a_removed_story() {
        let session = EditSession::new(21.0).unwrap();
        let doc = session.core.engine.doc();
        doc.create_story_with_paragraph_id("body", "p1", "ABCDE", "Normal", "left")
            .unwrap();
        session.layout_document_json(&caret_layout_input()).unwrap();
        session.build_display_list_frame("{}", 0.0).unwrap();
        session.set_selection("body", "p1", 2, "p1", 2).unwrap();
        let before: Value =
            serde_json::from_str(&session.resident_caret_snapshot_json().unwrap()).unwrap();
        assert!(before["caretRect"].is_object());
        let peer = EditingDoc::new(22);
        peer.apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        peer.delete_story("body").unwrap();
        doc.apply_update_v1(&peer.encode_state_as_update_v1())
            .unwrap();
        assert!(session.core.selection.borrow().is_some());
        let snapshot: Value =
            serde_json::from_str(&session.resident_caret_snapshot_json().unwrap()).unwrap();
        assert_eq!(snapshot["frameEpoch"], before["frameEpoch"]);
        assert!(snapshot["caretRect"].is_null());
    }

    #[test]
    fn resident_caret_snapshot_rebases_cumulative_block_embeds() {
        let session = EditSession::new(21.0).unwrap();
        seed_paragraph_after_embeds(
            session.core.engine.doc(),
            &["table", "pageBreak", "columnBreak", "blockSdt"],
            "ABCDE",
        );
        session.layout_document_json(&caret_layout_input()).unwrap();
        session.build_display_list_frame("{}", 0.0).unwrap();
        session.set_selection("body", "p1", 7, "p1", 7).unwrap();

        let snapshot: Value =
            serde_json::from_str(&session.resident_caret_snapshot_json().unwrap()).unwrap();
        let caret = &snapshot["caretRect"];

        assert_eq!(snapshot["frameEpoch"], 1);
        assert_eq!(caret["pageIndex"], 0);
        assert!((caret["x"].as_f64().unwrap() - 40.0).abs() < 0.001);
    }

    fn content_controls_docx() -> Vec<u8> {
        std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../packages/docx/src/yrs/__fixtures__/content-controls/template.docx"
        ))
        .unwrap()
    }

    fn control_text(session: &EditSession, tag: &str) -> Value {
        let listed = envelope(&session.list_content_controls_json("{}").unwrap());
        listed["content"]["controls"]
            .as_array()
            .unwrap()
            .iter()
            .find(|control| control["tag"] == tag)
            .map(|control| control["value"].clone())
            .unwrap()
    }

    #[test]
    fn content_control_envelopes_refuse_as_data() {
        let bytes = content_controls_docx();
        let session = EditSession::new(72.0).unwrap();
        session.core.open_docx(&bytes, true, None, None).unwrap();
        let listed = envelope(&session.list_content_controls_json("{}").unwrap());
        assert_eq!(listed["ok"], true);
        assert_eq!(listed["version"], session.version().as_str());
        assert_eq!(listed["content"]["schemaVersion"], 1);
        assert_eq!(listed["content"]["anchorScope"], "session");
        assert_eq!(listed["content"]["controls"].as_array().unwrap().len(), 7);
        assert_eq!(
            listed["content"]["controls"][0]["parentControlId"],
            Value::Null
        );
        let refused = envelope(
            &session
                .list_content_controls_json(r#"{"maxControls":1}"#)
                .unwrap(),
        );
        assert_eq!(refused["ok"], false);
        assert_eq!(refused["failure"]["code"], "limit-exceeded");
        let found = envelope(
            &session
                .find_content_controls_json(r#"{"kind":"alias","alias":"Address"}"#, "{}")
                .unwrap(),
        );
        assert_eq!(found["content"]["controls"][0]["controlId"], "body:sdt0");

        let request = json!({
            "expectVersion": session.version(),
            "steps": [{"op": "setContentControlText",
                "target": {"kind": "tag", "tag": "account.reference"}, "text": "x"}]
        });
        let ambiguous = envelope(&session.apply_edits_json(&request.to_string()).unwrap());
        assert_eq!(ambiguous["failure"]["code"], "ambiguous-target");
        assert_eq!(ambiguous["failure"]["reason"], "ambiguous-tag");

        session.clear_content_control_value("102").unwrap();
        assert_eq!(
            control_text(&session, "account.reference")["text"],
            "REF-000"
        );

        let snapshot = envelope(&list_docx_content_controls_json(&bytes, "{}").unwrap());
        assert_eq!(snapshot["ok"], true);
        assert_eq!(snapshot["content"]["anchorScope"], "snapshot");
        let matched = envelope(
            &find_docx_content_controls_json(
                &bytes,
                r#"{"kind":"tag","tag":"customer.name"}"#,
                "{}",
            )
            .unwrap(),
        );
        assert_eq!(
            matched["content"]["controls"][0]["showingPlaceholder"],
            true
        );
    }

    #[test]
    fn loading_shared_state_keeps_legacy_values_as_they_are() {
        let bytes = content_controls_docx();
        let author = EditSession::new(73.0).unwrap();
        author.core.open_docx(&bytes, true, None, None).unwrap();
        {
            let doc = author.core.engine.doc();
            let stories = doc.yrs_doc().transact().get_map(crate::STORIES).unwrap();
            let mut txn = doc.yrs_doc().transact_mut();
            let Some(yrs::Out::YText(body)) = stories.get(&txn, "body") else {
                panic!("body");
            };
            let map = yrs::Text::diff(&body, &txn, yrs::types::text::YChange::identity)
                .into_iter()
                .find_map(|diff| match diff.insert {
                    yrs::Out::YMap(map)
                        if map.get(&txn, "tag")
                            == Some(yrs::Out::Any(Any::from("account.reference"))) =>
                    {
                        Some(map)
                    }
                    _ => None,
                })
                .unwrap();
            map.insert(&mut txn, "value", "REF-LEGACY");
        }
        let joiner = EditSession::new(74.0).unwrap();
        joiner.core.open_docx(&bytes, false, None, None).unwrap();
        joiner.load(&author.encode_state()).unwrap();
        assert_eq!(
            control_text(&joiner, "account.reference")["text"],
            "REF-000"
        );
        assert_eq!(joiner.encode_state_vector(), author.encode_state_vector());
        let listed = envelope(&joiner.list_content_controls_json("{}").unwrap());
        assert!(
            listed["content"]["diagnostics"]
                .as_array()
                .unwrap()
                .iter()
                .any(|diagnostic| diagnostic["code"] == "legacy-control-value")
        );
        assert!(!joiner.can_undo());
    }

    #[test]
    fn unpaired_surrogates_in_fill_text_are_invalid_text() {
        let session = EditSession::new(75.0).unwrap();
        session
            .core
            .open_docx(&content_controls_docx(), true, None, None)
            .unwrap();
        let request = |text: &str| {
            format!(
                r#"{{"expectVersion":"{}","steps":[{{"op":"setContentControlText","target":{{"kind":"tag","tag":"document.title"}},"text":"{text}"}}]}}"#,
                session.version()
            )
        };
        for outcome in [
            session.validate_edits_json(&request(r"a\ud800b")).unwrap(),
            session.apply_edits_json(&request(r"\udc00")).unwrap(),
        ] {
            let outcome = envelope(&outcome);
            assert_eq!(outcome["ok"], false);
            assert_eq!(outcome["failure"]["code"], "invalid-step");
            assert_eq!(outcome["failure"]["reason"], "invalid-text");
            assert_eq!(outcome["failure"]["stepIndex"], 0);
        }
        assert!(!session.can_undo());
        let paired = envelope(
            &session
                .validate_edits_json(&request(r"\ud83d\ude00"))
                .unwrap(),
        );
        assert_eq!(paired["wouldApply"], true);
        let elsewhere = format!(
            r#"{{"expectVersion":"{}","steps":[{{"op":"setContentControlText","target":{{"kind":"tag","tag":"\ud800"}},"text":"x"}}]}}"#,
            session.version()
        );
        assert!(EditRequest::from_json(&elsewhere).is_err());
    }

    fn compare_fixture(body: &str) -> Vec<u8> {
        ooxml_opc::rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec()),
            ("word/document.xml".to_owned(), format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr/></w:body></w:document>"#).into_bytes()),
        ])
        .unwrap()
    }

    #[test]
    fn compare_bridge_applies_changes_and_verifies_saved_bytes() {
        let original = compare_fixture("<w:p><w:r><w:t>old text</w:t></w:r></w:p>");
        let revised = compare_fixture("<w:p><w:r><w:t>new text</w:t></w:r></w:p>");
        let options = r#"{"author":"A","date":"2024-01-02T03:04:05Z"}"#;
        let session = EditSession::new(1.0).unwrap();
        let outcome: Value = serde_json::from_str(
            &session
                .compare_docx_json(&original, &revised, options)
                .unwrap(),
        )
        .unwrap();
        assert_eq!(outcome["ok"], true);
        assert_eq!(outcome["noop"], false);
        assert!(outcome.get("changes").is_none());
        assert_eq!(
            outcome["save"]["paragraphs"],
            json!([{ "path": [0, 0], "block": 0 }])
        );
        assert_eq!(outcome["save"]["now"], "2024-01-02T03:04:05.000Z");
        assert!(session.materialize_docx().unwrap().is_some());
        let unchanged: Value =
            serde_json::from_str(&session.finish_compared_docx_json(&original).unwrap()).unwrap();
        assert_eq!(unchanged["ok"], false);
        assert_eq!(unchanged["diagnostics"][0]["code"], "roundtrip-mismatch");

        let failing = EditSession::new(4.0).unwrap();
        failing
            .compare_docx_json(&original, &revised, options)
            .unwrap();
        let failed: Value =
            serde_json::from_str(&failing.fail_compared_docx_json("disk full").unwrap()).unwrap();
        assert_eq!(failed["ok"], false);
        assert_eq!(failed["diagnostics"][0]["code"], "serialization-failed");

        let fresh = EditSession::new(2.0).unwrap();
        let noop: Value = serde_json::from_str(
            &fresh
                .compare_docx_json(&original, &original, options)
                .unwrap(),
        )
        .unwrap();
        assert_eq!(
            noop,
            json!({ "ok": true, "noop": true, "changes": [], "diagnostics": [] })
        );
        let refused: Value = serde_json::from_str(
            &EditSession::new(3.0)
                .unwrap()
                .compare_docx_json(
                    &original,
                    &revised,
                    r#"{"author":"","date":"2024-01-02T03:04:05Z"}"#,
                )
                .unwrap(),
        )
        .unwrap();
        assert_eq!(refused["ok"], false);
        assert_eq!(refused["diagnostics"][0]["code"], "invalid-options");
    }

    /// A session holding two body paragraphs and a header story whose body is
    /// laid out and shown, so resident input is ready for it.
    fn shown_text_session(client_id: f64) -> EditSession {
        let session = EditSession::new(client_id).unwrap();
        let doc = session.core.engine.doc();
        doc.create_story("body", "Alpha beta", "Normal", "left")
            .unwrap();
        doc.split_paragraph(&EditCtx::local("", ""), Position::new("body", 5), None)
            .unwrap();
        doc.create_story("hf:rId1", "Header", "Normal", "left")
            .unwrap();
        let font = session
            .register_measure_font(include_bytes!(
                "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
            ))
            .unwrap();
        let mut request = json!({
            "bodyStory": "body",
            "regions": {"sections": [{"properties": {}}]},
            "renderEnv": {},
        });
        let requirements = envelope(
            &session
                .layout_font_requirements_json(&request.to_string())
                .unwrap(),
        );
        let chains: serde_json::Map<String, Value> = requirements
            .as_array()
            .unwrap()
            .iter()
            .map(|requirement| {
                (
                    requirement["key"].as_str().unwrap().to_owned(),
                    json!([font]),
                )
            })
            .collect();
        request["measurement"] = json!({
            "fontChains": chains,
            "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
        });
        session
            .layout_document_with_regions_retained(&request.to_string())
            .unwrap();
        session
            .build_display_list_frame(&json!({ "fontChains": chains }).to_string(), 0.0)
            .unwrap();
        session
    }

    #[test]
    fn text_commands_take_the_generic_path_where_resident_input_is_not_ready() {
        let resident = shown_text_session(78.0);
        let generic = shown_text_session(78.0);
        let paragraphs = resident.core.engine.doc().paragraphs("body").unwrap();
        let (first, second) = (&paragraphs[0].para_id, &paragraphs[1].para_id);
        let header = &resident.core.engine.doc().paragraphs("hf:rId1").unwrap()[0].para_id;
        resident
            .edit_text_resident("body", first, 1, first, 1, Some("x".to_owned()))
            .unwrap();
        assert!(resident.core.resident_text_edit.borrow().is_some());
        generic
            .insert_text("body", first, 1, "x", None, None)
            .unwrap();
        assert_eq!(
            resident
                .edit_text_resident("body", first, 2, second, 1, None)
                .unwrap(),
            generic
                .delete_range("body", first, 2, second, 1, None, None)
                .unwrap()
        );
        assert!(resident.core.resident_text_edit.borrow().is_none());
        assert_eq!(
            resident
                .edit_text_resident("hf:rId1", header, 1, header, 1, Some("y".to_owned()))
                .unwrap(),
            generic
                .insert_text("hf:rId1", header, 1, "y", None, None)
                .unwrap()
        );
        assert!(resident.core.resident_text_edit.borrow().is_none());
        for story in ["body", "hf:rId1"] {
            assert_eq!(
                resident.core.engine.doc().story_segments(story).unwrap(),
                generic.core.engine.doc().story_segments(story).unwrap(),
                "{story}"
            );
        }
    }

    #[test]
    fn an_empty_text_command_records_no_resident_edit() {
        let session = shown_text_session(79.0);
        let paragraph = session.core.engine.doc().paragraphs("body").unwrap()[0]
            .para_id
            .clone();
        let before = session.core.engine.doc().story_segments("body").unwrap();
        session
            .edit_text_resident("body", &paragraph, 1, &paragraph, 1, Some(String::new()))
            .unwrap();
        assert!(session.core.resident_text_edit.borrow().is_none());
        assert_eq!(
            session.core.engine.doc().story_segments("body").unwrap(),
            before
        );
    }
}
