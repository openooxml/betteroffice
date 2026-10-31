//! the single wasm-bindgen boundary: coarse json-string calls over the pure
//! `Session` methods in `core.rs`.

use wasm_bindgen::prelude::*;
use xlsx_engine::{Document, EngineError, ErrorKind, Host, SnapshotBuilder};

#[wasm_bindgen(typescript_custom_section)]
const DECODED_FONT_FACE: &str =
    r#"export interface DecodedFontFace { font: Uint8Array; payload?: Uint8Array; }"#;

/// Decodes one font pack face (codec "zstd" or "woff2t+zstd") with this engine's linked decoder.
/// Returns owned bytes and an optional transformed payload. Non-zstd input passes through.
/// The default output limit is 64 MiB; woff2t decoding is always capped at 64 MiB.
#[wasm_bindgen(js_name = decodeFontFace, unchecked_return_type = "DecodedFontFace")]
pub fn decode_font_face(
    bytes: &[u8],
    codec: &str,
    #[wasm_bindgen(unchecked_optional_param_type = "Uint8Array | null")] dict: Option<Vec<u8>>,
    #[wasm_bindgen(js_name = maxOutput, unchecked_optional_param_type = "number | null")]
    max_output: Option<f64>,
) -> Result<JsValue, JsError> {
    let decoded =
        betteroffice_fonts::decode_font_face_js(bytes, codec, dict.as_deref(), max_output)
            .map_err(|error| JsError::new(&error.to_string()))?;
    let object = js_sys::Object::new();
    js_sys::Reflect::set(
        &object,
        &"font".into(),
        &js_sys::Uint8Array::from(decoded.font.as_slice()),
    )
    .map_err(|_| JsError::new("cannot create decoded font face"))?;
    if let Some(payload) = decoded.payload {
        js_sys::Reflect::set(
            &object,
            &"payload".into(),
            &js_sys::Uint8Array::from(payload.as_slice()),
        )
        .map_err(|_| JsError::new("cannot create decoded font face"))?;
    }
    Ok(object.into())
}

#[cfg(all(target_arch = "wasm32", not(feature = "raster")))]
#[used]
#[unsafe(link_section = "betteroffice.xlsx.no-png")]
static PNG_EXPORT_UNAVAILABLE: [u8; 1] = [1];

/// days from the 1900 epoch to 1970-01-01, phantom-1900 leap day included.
const UNIX_EPOCH_SERIAL: f64 = 25569.0;
const MS_PER_DAY: f64 = 86_400_000.0;

/// current local time as a 1900-system serial for volatile cells; computed only
/// at the wasm boundary so the pure core stays deterministic and native-testable.
fn now_serial() -> Result<Option<f64>, EngineError> {
    let ms = js_sys::Date::now();
    let offset = js_sys::Date::new(&ms.into()).get_timezone_offset();
    Ok(Some(
        (ms - offset * 60_000.0) / MS_PER_DAY + UNIX_EPOCH_SERIAL,
    ))
}

#[wasm_bindgen]
extern "C" {
    #[wasm_bindgen(js_namespace = performance, js_name = now)]
    fn performance_now() -> f64;
}

#[cfg(target_arch = "wasm32")]
#[wasm_bindgen]
extern "C" {
    #[wasm_bindgen(js_namespace = console, js_name = error)]
    fn console_error(message: &str);
}

/// wasm panics abort, so a trap reaches JS as a bare `RuntimeError:
/// unreachable executed`. Log the panic first so it stays diagnosable.
#[cfg(target_arch = "wasm32")]
#[wasm_bindgen(start)]
pub fn install_panic_hook() {
    static INSTALLED: std::sync::Once = std::sync::Once::new();
    INSTALLED.call_once(|| {
        std::panic::set_hook(Box::new(|info| {
            console_error(&format!("[xlsx-wasm] panic: {info}"));
        }));
    });
}

/// records the size of an allocation the engine could not get: wasm aborts on
/// allocation failure, so the trap that follows is told apart from a panic.
#[cfg(target_arch = "wasm32")]
mod allocator {
    use std::alloc::{GlobalAlloc, Layout, System};
    use std::sync::atomic::{AtomicUsize, Ordering};

    pub(crate) static FAILED: AtomicUsize = AtomicUsize::new(0);

    struct Recording;

    fn record(ptr: *mut u8, size: usize) -> *mut u8 {
        if ptr.is_null() {
            FAILED.store(size.max(1), Ordering::Relaxed);
        }
        ptr
    }

    unsafe impl GlobalAlloc for Recording {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            record(unsafe { System.alloc(layout) }, layout.size())
        }

        unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
            record(unsafe { System.alloc_zeroed(layout) }, layout.size())
        }

        unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
            unsafe { System.dealloc(ptr, layout) }
        }

        unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
            record(unsafe { System.realloc(ptr, layout, new_size) }, new_size)
        }
    }

    #[global_allocator]
    static ALLOCATOR: Recording = Recording;
}

/// bytes of the last allocation the engine could not get, or 0.
#[cfg(target_arch = "wasm32")]
#[wasm_bindgen(js_name = allocationFailure)]
pub fn allocation_failure() -> f64 {
    allocator::FAILED.load(std::sync::atomic::Ordering::Relaxed) as f64
}

/// tells the worker host that a long engine call is still running, through
/// `globalThis.__betterofficeEngineProgress` when the host installed one.
fn engine_progress() {
    let hook = js_sys::Reflect::get(&js_sys::global(), &"__betterofficeEngineProgress".into());
    if let Ok(hook) = hook.and_then(|hook| hook.dyn_into::<js_sys::Function>()) {
        let _ = hook.call0(&JsValue::UNDEFINED);
    }
}

fn wasm_host() -> Host {
    Host {
        now_serial,
        timer_ms: performance_now,
        progress: engine_progress,
        version: xlsx_engine::Session::version(),
    }
}

fn map_engine_error(error: EngineError) -> JsValue {
    match error.kind {
        ErrorKind::Message => JsValue::from_str(&error.message),
        ErrorKind::Error => js_sys::Error::new(&error.message).into(),
        ErrorKind::TypeError => js_sys::TypeError::new(&error.message).into(),
        ErrorKind::PendingException => unreachable!(),
    }
}

/// a workbook handle exposed to js; wraps the pure `Session`.
#[wasm_bindgen]
pub struct XlsxDocument {
    inner: Document,
}

#[wasm_bindgen]
pub struct XlsxSnapshotBuilder {
    inner: SnapshotBuilder,
}

#[wasm_bindgen]
impl XlsxSnapshotBuilder {
    #[wasm_bindgen(constructor)]
    pub fn new() -> XlsxSnapshotBuilder {
        Self {
            inner: SnapshotBuilder::new(wasm_host()),
        }
    }

    pub fn push(&mut self, chunk: &[u8]) -> Result<(), JsValue> {
        self.inner.push(chunk).map_err(map_engine_error)
    }

    pub fn advance(&mut self, records: usize, bytes: usize) -> Result<bool, JsValue> {
        self.inner.advance(records, bytes).map_err(map_engine_error)
    }

    pub fn finish(self) -> Result<XlsxDocument, JsValue> {
        self.inner
            .finish()
            .map(|inner| XlsxDocument { inner })
            .map_err(map_engine_error)
    }
}

impl Default for XlsxSnapshotBuilder {
    fn default() -> Self {
        Self::new()
    }
}

#[wasm_bindgen]
impl XlsxDocument {
    #[wasm_bindgen(js_name = recoveryCheckpoint)]
    pub fn recovery_checkpoint(
        &self,
        through_seq: f64,
        preserve_history: Option<bool>,
    ) -> Result<Vec<u8>, JsValue> {
        self.inner
            .recovery_checkpoint(through_seq, preserve_history)
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = openRecoveryCheckpoint)]
    pub fn open_recovery_checkpoint(
        bytes: &[u8],
        through_seq: f64,
    ) -> Result<XlsxDocument, JsValue> {
        Document::open_recovery_checkpoint(bytes, through_seq, wasm_host())
            .map(|inner| XlsxDocument { inner })
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = beginPeerSnapshot)]
    pub fn begin_peer_snapshot(&mut self, records: usize, bytes: usize) -> Result<(), JsValue> {
        self.inner
            .begin_peer_snapshot(records, bytes)
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = nextPeerSnapshotChunk)]
    pub fn next_peer_snapshot_chunk(&mut self) -> Result<Option<js_sys::Uint8Array>, JsValue> {
        self.inner
            .next_peer_snapshot_chunk()
            .map(|chunk| chunk.map(|bytes| js_sys::Uint8Array::from(bytes.as_slice())))
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = endPeerSnapshot)]
    pub fn end_peer_snapshot(&mut self) {
        self.inner.end_peer_snapshot();
    }

    #[doc(hidden)]
    #[wasm_bindgen(js_name = adoptPeerVersion)]
    pub fn adopt_peer_version(&mut self, version: &str) -> Result<(), JsValue> {
        self.inner
            .adopt_peer_version(version)
            .map_err(map_engine_error)
    }

    #[doc(hidden)]
    #[wasm_bindgen(js_name = peerHydrationJson)]
    pub fn peer_hydration_json(&self) -> Result<String, JsValue> {
        self.inner.peer_hydration_json().map_err(map_engine_error)
    }

    #[doc(hidden)]
    #[wasm_bindgen(js_name = openWithPeerHydrationJson)]
    pub fn open_with_peer_hydration_json(
        bytes: &[u8],
        hydration: &str,
    ) -> Result<XlsxDocument, JsValue> {
        Document::open_with_peer_hydration_json(bytes, hydration, wasm_host())
            .map(|inner| XlsxDocument { inner })
            .map_err(map_engine_error)
    }

    /// open a workbook from raw `.xlsx` bytes.
    pub fn open(bytes: &[u8]) -> Result<XlsxDocument, JsValue> {
        Document::open(bytes, wasm_host())
            .map(|inner| XlsxDocument { inner })
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = openWithCalculationJson)]
    pub fn open_with_calculation_json(
        bytes: &[u8],
        context: &str,
    ) -> Result<XlsxDocument, JsValue> {
        Document::open_with_calculation_json(bytes, context, wasm_host())
            .map(|inner| XlsxDocument { inner })
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = setCalculationContextJson)]
    pub fn set_calculation_context_json(&mut self, context: &str) -> Result<(), JsValue> {
        self.inner
            .set_calculation_context_json(context)
            .map_err(map_engine_error)
    }

    /// Open a replica with a positive, safe-integer client ID.
    #[wasm_bindgen(js_name = openCollaborative)]
    pub fn open_collaborative(bytes: &[u8], client_id: f64) -> Result<XlsxDocument, JsValue> {
        Document::open_collaborative(bytes, client_id, wasm_host())
            .map(|inner| XlsxDocument { inner })
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(getter, js_name = clientId)]
    pub fn client_id(&self) -> f64 {
        self.inner.client_id()
    }

    #[wasm_bindgen(js_name = encodeStateVector)]
    pub fn encode_state_vector(&self) -> Vec<u8> {
        self.inner.encode_state_vector()
    }

    #[wasm_bindgen(js_name = encodeStateAsUpdate)]
    pub fn encode_state_as_update(&self) -> Vec<u8> {
        self.inner.encode_state_as_update()
    }

    #[wasm_bindgen(js_name = encodeDiff)]
    pub fn encode_diff(&self, remote_state_vector: &[u8]) -> Result<Vec<u8>, JsValue> {
        self.inner
            .encode_diff(remote_state_vector)
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = applyUpdateJson)]
    pub fn apply_update_json(&mut self, update: &[u8]) -> Result<String, JsValue> {
        self.inner
            .apply_update_json(update)
            .map_err(map_engine_error)
    }

    /// Start queuing origin-prefixed Yrs update events for polling.
    #[wasm_bindgen(js_name = startUpdateObservation)]
    pub fn start_update_observation(&mut self) -> Result<(), JsValue> {
        self.inner
            .start_update_observation()
            .map_err(map_engine_error)
    }

    /// Stop observation and discard queued events.
    #[wasm_bindgen(js_name = clearUpdateObservation)]
    pub fn clear_update_observation(&mut self) {
        self.inner.clear_update_observation();
    }

    /// Poll one event: origin byte (`0` local, `1` remote), then update; empty means none.
    #[wasm_bindgen(js_name = drainUpdateEvent)]
    pub fn drain_update_event(&self) -> Result<Vec<u8>, JsValue> {
        self.inner.drain_update_event().map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = printDisplayListJson)]
    pub fn print_display_list_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner
            .print_display_list_json(args)
            .map_err(map_engine_error)
    }

    /// serialized `DisplayList` for a serialized `Viewport`.
    #[wasm_bindgen(js_name = displayListJson)]
    pub fn display_list_json(&self, viewport_json: &str) -> Result<String, JsValue> {
        self.inner
            .display_list_json(viewport_json)
            .map_err(map_engine_error)
    }

    /// @experimental
    #[wasm_bindgen(js_name = displayListForJson)]
    pub fn display_list_for_json(
        &self,
        sheet: u32,
        viewport_json: &str,
    ) -> Result<String, JsValue> {
        self.inner
            .display_list_for_json(sheet, viewport_json)
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = displayListProfiledJson)]
    pub fn display_list_profiled_json(&self, viewport_json: &str) -> Result<String, JsValue> {
        self.inner
            .display_list_profiled_json(viewport_json)
            .map_err(map_engine_error)
    }

    /// the chart under a viewport-local point, or `null`.
    #[wasm_bindgen(js_name = chartAtPointJson)]
    pub fn chart_at_point_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner
            .chart_at_point_json(args)
            .map_err(map_engine_error)
    }

    /// slide a chart by a pixel delta as one undo step.
    #[wasm_bindgen(js_name = moveChartJson)]
    pub fn move_chart_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.inner.move_chart_json(args).map_err(map_engine_error)
    }

    /// render the current sheet viewport to png bytes (raster feature only).
    #[wasm_bindgen(js_name = renderPng)]
    pub fn render_png(&self, viewport_json: &str) -> Result<Vec<u8>, JsValue> {
        #[cfg(feature = "raster")]
        {
            self.inner
                .render_png(viewport_json)
                .map_err(map_engine_error)
        }
        #[cfg(not(feature = "raster"))]
        {
            let _ = viewport_json;
            Err(JsValue::from_str("png export not in this build"))
        }
    }

    /// render an a1 range (default: used range) at an optional scale to png.
    #[wasm_bindgen(js_name = renderRangePng)]
    pub fn render_range_png(&self, args: &str) -> Result<Vec<u8>, JsValue> {
        #[cfg(feature = "raster")]
        {
            self.inner.render_range_png(args).map_err(map_engine_error)
        }
        #[cfg(not(feature = "raster"))]
        {
            let _ = args;
            Err(JsValue::from_str("png export not in this build"))
        }
    }

    /// serialized `SheetInfo`: stable IDs, names, active index, content extent.
    #[wasm_bindgen(js_name = sheetInfoJson)]
    pub fn sheet_info_json(&self) -> Result<String, JsValue> {
        self.inner.sheet_info_json().map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = sheetCount)]
    pub fn sheet_count(&self) -> usize {
        self.inner.sheet_count()
    }

    #[wasm_bindgen(js_name = sheetInfoForJson)]
    pub fn sheet_info_for_json(&self, sheet: u32) -> Result<String, JsValue> {
        self.inner
            .sheet_info_for_json(sheet)
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = selectionDataJson)]
    pub fn selection_data_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner
            .selection_data_json(args)
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = cellRectJson)]
    pub fn cell_rect_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.cell_rect_json(args).map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = calculationStatusJson)]
    pub fn calculation_status_json(&self) -> Result<String, JsValue> {
        self.inner
            .calculation_status_json()
            .map_err(map_engine_error)
    }

    /// switch the active sheet by index.
    #[wasm_bindgen(js_name = setActiveSheet)]
    pub fn set_active_sheet(&mut self, index: u32) -> Result<(), JsValue> {
        self.inner.set_active_sheet(index).map_err(map_engine_error)
    }

    /// enter one cell edit; returns updated `SheetInfo` json.
    #[wasm_bindgen(js_name = editCellJson)]
    pub fn edit_cell_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.inner.edit_cell_json(args).map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = editCellProfiledJson)]
    pub fn edit_cell_profiled_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.inner
            .edit_cell_profiled_json(args)
            .map_err(map_engine_error)
    }

    /// enter a batch of cell edits as one undo step; returns `SheetInfo` json.
    #[wasm_bindgen(js_name = editCellsJson)]
    pub fn edit_cells_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.inner.edit_cells_json(args).map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = moveRangeJson)]
    pub fn move_range_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.inner.move_range_json(args).map_err(map_engine_error)
    }

    /// apply a raw op list as one user transaction; returns `SheetInfo` json.
    #[wasm_bindgen(js_name = applyOpsJson)]
    pub fn apply_ops_json(&mut self, transaction_json: &str) -> Result<String, JsValue> {
        self.inner
            .apply_ops_json(transaction_json)
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = chartsJson)]
    pub fn charts_json(&self, sheet: u32) -> Result<String, JsValue> {
        self.inner.charts_json(sheet).map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = editChartsJson)]
    pub fn edit_charts_json(&mut self, edits: &str) -> Result<String, JsValue> {
        self.inner.edit_charts_json(edits).map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = conditionalFormatsJson)]
    pub fn conditional_formats_json(&self, sheet: u32) -> Result<String, JsValue> {
        self.inner
            .conditional_formats_json(sheet)
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = applyOpsProfiledJson)]
    pub fn apply_ops_profiled_json(&mut self, transaction_json: &str) -> Result<String, JsValue> {
        self.inner
            .apply_ops_profiled_json(transaction_json)
            .map_err(map_engine_error)
    }

    /// undo the last transaction; returns `{"applied":bool,"sheetInfo":{...}}`.
    #[wasm_bindgen(js_name = undoJson)]
    pub fn undo_json(&mut self) -> Result<String, JsValue> {
        self.inner.undo_json().map_err(map_engine_error)
    }

    /// redo the last undone transaction; same shape as `undoJson`.
    #[wasm_bindgen(js_name = redoJson)]
    pub fn redo_json(&mut self) -> Result<String, JsValue> {
        self.inner.redo_json().map_err(map_engine_error)
    }

    /// the editable representation of one cell.
    #[wasm_bindgen(js_name = cellJson)]
    pub fn cell_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.cell_json(args).map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = searchTextJson)]
    pub fn search_text_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.search_text_json(args).map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = cellPositionJson)]
    pub fn cell_position_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner
            .cell_position_json(args)
            .map_err(map_engine_error)
    }

    /// a rectangular block of cells for clipboard copy.
    #[wasm_bindgen(js_name = rangeCellsJson)]
    pub fn range_cells_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.range_cells_json(args).map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = patchRangeStyleJson)]
    pub fn patch_range_style_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.inner
            .patch_range_style_json(args)
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = setRangeNumberFormatJson)]
    pub fn set_range_number_format_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.inner
            .set_range_number_format_json(args)
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = selectionFormattingJson)]
    pub fn selection_formatting_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner
            .selection_formatting_json(args)
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = captureFormatJson)]
    pub fn capture_format_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner
            .capture_format_json(args)
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = applyFormatJson)]
    pub fn apply_format_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.inner.apply_format_json(args).map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = mergedRangesJson)]
    pub fn merged_ranges_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner
            .merged_ranges_json(args)
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = visibleMergedRangesJson)]
    pub fn visible_merged_ranges_json(
        &self,
        sheet: u32,
        viewport_json: &str,
    ) -> Result<String, JsValue> {
        self.inner
            .visible_merged_ranges_json(sheet, viewport_json)
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = historyStateJson)]
    pub fn history_state_json(&self) -> Result<String, JsValue> {
        self.inner.history_state_json().map_err(map_engine_error)
    }

    /// register an agent proposal (preview only); returns the stored `Proposal` json.
    #[wasm_bindgen(js_name = proposeJson)]
    pub fn propose_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.inner.propose_json(args).map_err(map_engine_error)
    }

    /// the pending proposals: `{"proposals":[...]}`.
    #[wasm_bindgen(js_name = listProposalsJson)]
    pub fn list_proposals_json(&self) -> Result<String, JsValue> {
        self.inner.list_proposals_json().map_err(map_engine_error)
    }

    /// accept a proposal as one agent transaction; returns the edit envelope
    /// plus `proposalId`, or a `{"code":"staleProposal",...}` error when the base moved.
    #[wasm_bindgen(js_name = acceptProposalJson)]
    pub fn accept_proposal_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.inner
            .accept_proposal_json(args)
            .map_err(map_engine_error)
    }

    /// reject a proposal by id; returns `{"removed":bool}`.
    #[wasm_bindgen(js_name = rejectProposalJson)]
    pub fn reject_proposal_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.inner
            .reject_proposal_json(args)
            .map_err(map_engine_error)
    }

    /// The session-scoped version of the committed workbook state.
    #[wasm_bindgen(js_name = documentVersion)]
    pub fn document_version(&self) -> String {
        self.inner.document_version()
    }

    #[wasm_bindgen(js_name = readCellsJson)]
    pub fn read_cells_json(&self, request: &str) -> Result<String, JsValue> {
        self.inner
            .read_cells_json(request)
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = listDefinedNamesJson)]
    pub fn list_defined_names_json(&self, request: &str) -> Result<String, JsValue> {
        self.inner
            .list_defined_names_json(request)
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = findTextJson)]
    pub fn find_text_json(&self, request: &str) -> Result<String, JsValue> {
        self.inner.find_text_json(request).map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = validateEditsJson)]
    pub fn validate_edits_json(&self, request: &str) -> Result<String, JsValue> {
        self.inner
            .validate_edits_json(request)
            .map_err(map_engine_error)
    }

    /// apply an edit batch all-or-nothing; volatile functions see only the request's clock.
    #[wasm_bindgen(js_name = applyEditsJson)]
    pub fn apply_edits_json(&mut self, request: &str) -> Result<String, JsValue> {
        self.inner
            .apply_edits_json(request)
            .map_err(map_engine_error)
    }

    /// Export the committed workbook with the version it was read at; nothing recalculates.
    #[wasm_bindgen(js_name = exportStructuredJson)]
    pub fn export_structured_json(&self, options: &str) -> Result<String, JsValue> {
        self.inner
            .export_structured_json(options)
            .map_err(map_engine_error)
    }

    #[wasm_bindgen(js_name = exportMarkdownJson)]
    pub fn export_markdown_json(
        &self,
        options: &str,
        markdown_options: &str,
    ) -> Result<String, JsValue> {
        self.inner
            .export_markdown_json(options, markdown_options)
            .map_err(map_engine_error)
    }

    /// serialize the current workbook back to `.xlsx` bytes.
    #[wasm_bindgen(js_name = saveBytes)]
    pub fn save_bytes(&self) -> Result<Vec<u8>, JsValue> {
        self.inner.save_bytes().map_err(map_engine_error)
    }

    /// crate version string.
    pub fn version() -> String {
        xlsx_engine::Session::version().to_string()
    }
}

#[wasm_bindgen(js_name = displayListCellLimit)]
pub fn display_list_cell_limit() -> f64 {
    xlsx_engine::display_list_cell_limit()
}

/// Export `.xlsx` bytes as read: stored formula results, no clock.
#[wasm_bindgen(js_name = exportXlsxStructuredJson)]
pub fn export_xlsx_structured_json(bytes: &[u8], options: &str) -> Result<String, JsValue> {
    xlsx_engine::export_xlsx_structured_json(bytes, options).map_err(map_engine_error)
}

#[wasm_bindgen(js_name = exportXlsxMarkdownJson)]
pub fn export_xlsx_markdown_json(
    bytes: &[u8],
    options: &str,
    markdown_options: &str,
) -> Result<String, JsValue> {
    xlsx_engine::export_xlsx_markdown_json(bytes, options, markdown_options)
        .map_err(map_engine_error)
}

#[wasm_bindgen(js_name = renderXlsxMarkdownJson)]
pub fn render_xlsx_markdown_json(content: &str, options: &str) -> Result<String, JsValue> {
    xlsx_engine::render_xlsx_markdown_json(content, options).map_err(map_engine_error)
}
