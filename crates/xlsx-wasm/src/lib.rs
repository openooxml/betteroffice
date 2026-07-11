//! the single wasm-bindgen boundary: coarse json-string calls only. every
//! exported method is a one-line wrapper over a pure `Session` method in `core.rs`.

mod core;

use wasm_bindgen::prelude::*;

use crate::core::Session;

/// days from the 1900 epoch to 1970-01-01, phantom-1900 leap day included.
const UNIX_EPOCH_SERIAL: f64 = 25569.0;
const MS_PER_DAY: f64 = 86_400_000.0;

/// current utc time as a 1900-system serial for volatile cells; computed only
/// at the wasm boundary so the pure core stays deterministic and native-testable.
fn now_serial() -> Option<f64> {
    Some(js_sys::Date::now() / MS_PER_DAY + UNIX_EPOCH_SERIAL)
}

/// a workbook handle exposed to js; wraps the pure `Session`.
#[wasm_bindgen]
pub struct XlsxDocument {
    session: Session,
}

#[wasm_bindgen]
impl XlsxDocument {
    /// open a workbook from raw `.xlsx` bytes.
    pub fn open(bytes: &[u8]) -> Result<XlsxDocument, JsValue> {
        Session::open(bytes, now_serial())
            .map(|session| XlsxDocument { session })
            .map_err(|e| JsValue::from_str(&e))
    }

    /// serialized `DisplayList` for a serialized `Viewport`.
    #[wasm_bindgen(js_name = displayListJson)]
    pub fn display_list_json(&self, viewport_json: &str) -> Result<String, JsValue> {
        self.session
            .display_list_json(viewport_json)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// render the current sheet viewport to png bytes (raster feature only).
    #[cfg(feature = "raster")]
    #[wasm_bindgen(js_name = renderPng)]
    pub fn render_png(&self, viewport_json: &str) -> Result<Vec<u8>, JsValue> {
        self.session
            .render_png(viewport_json)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// render an a1 range (default: used range) at an optional scale to png.
    #[cfg(feature = "raster")]
    #[wasm_bindgen(js_name = renderRangePng)]
    pub fn render_range_png(&self, args: &str) -> Result<Vec<u8>, JsValue> {
        self.session
            .render_range_png(args)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// serialized `SheetInfo`: sheet names, active index, content extent.
    #[wasm_bindgen(js_name = sheetInfoJson)]
    pub fn sheet_info_json(&self) -> Result<String, JsValue> {
        self.session
            .sheet_info_json()
            .map_err(|e| JsValue::from_str(&e))
    }

    /// switch the active sheet by index.
    #[wasm_bindgen(js_name = setActiveSheet)]
    pub fn set_active_sheet(&mut self, index: u32) -> Result<(), JsValue> {
        self.session
            .set_active_sheet(index)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// enter one cell edit; returns updated `SheetInfo` json.
    #[wasm_bindgen(js_name = editCellJson)]
    pub fn edit_cell_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.session
            .edit_cell_json(args, now_serial())
            .map_err(|e| JsValue::from_str(&e))
    }

    /// enter a batch of cell edits as one undo step; returns `SheetInfo` json.
    #[wasm_bindgen(js_name = editCellsJson)]
    pub fn edit_cells_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.session
            .edit_cells_json(args, now_serial())
            .map_err(|e| JsValue::from_str(&e))
    }

    /// apply a raw op list as one user transaction; returns `SheetInfo` json.
    #[wasm_bindgen(js_name = applyOpsJson)]
    pub fn apply_ops_json(&mut self, transaction_json: &str) -> Result<String, JsValue> {
        self.session
            .apply_ops_json(transaction_json, now_serial())
            .map_err(|e| JsValue::from_str(&e))
    }

    /// undo the last transaction; returns `{"applied":bool,"sheetInfo":{...}}`.
    #[wasm_bindgen(js_name = undoJson)]
    pub fn undo_json(&mut self) -> Result<String, JsValue> {
        self.session
            .undo_json(now_serial())
            .map_err(|e| JsValue::from_str(&e))
    }

    /// redo the last undone transaction; same shape as `undoJson`.
    #[wasm_bindgen(js_name = redoJson)]
    pub fn redo_json(&mut self) -> Result<String, JsValue> {
        self.session
            .redo_json(now_serial())
            .map_err(|e| JsValue::from_str(&e))
    }

    /// the editable representation of one cell.
    #[wasm_bindgen(js_name = cellJson)]
    pub fn cell_json(&self, args: &str) -> Result<String, JsValue> {
        self.session
            .cell_json(args)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// a rectangular block of cells for clipboard copy.
    #[wasm_bindgen(js_name = rangeCellsJson)]
    pub fn range_cells_json(&self, args: &str) -> Result<String, JsValue> {
        self.session
            .range_cells_json(args)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// register an agent proposal (preview only); returns the stored `Proposal` json.
    #[wasm_bindgen(js_name = proposeJson)]
    pub fn propose_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.session
            .propose_json(args, now_serial())
            .map_err(|e| JsValue::from_str(&e))
    }

    /// the pending proposals: `{"proposals":[...]}`.
    #[wasm_bindgen(js_name = listProposalsJson)]
    pub fn list_proposals_json(&self) -> Result<String, JsValue> {
        self.session
            .list_proposals_json()
            .map_err(|e| JsValue::from_str(&e))
    }

    /// accept a proposal as one agent transaction; returns the edit envelope
    /// plus `proposalId`, or a `stale: ...` error when the base moved.
    #[wasm_bindgen(js_name = acceptProposalJson)]
    pub fn accept_proposal_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.session
            .accept_proposal_json(args, now_serial())
            .map_err(|e| JsValue::from_str(&e))
    }

    /// reject a proposal by id; returns `{"removed":bool}`.
    #[wasm_bindgen(js_name = rejectProposalJson)]
    pub fn reject_proposal_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.session
            .reject_proposal_json(args)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// serialize the current workbook back to `.xlsx` bytes.
    #[wasm_bindgen(js_name = saveBytes)]
    pub fn save_bytes(&self) -> Result<Vec<u8>, JsValue> {
        self.session.save().map_err(|e| JsValue::from_str(&e))
    }

    /// crate version string.
    pub fn version() -> String {
        Session::version().to_string()
    }
}
