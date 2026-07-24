//! the single wasm-bindgen boundary: coarse json-string calls over the pure
//! `Session` methods in `core.rs`.

mod core;

use std::collections::VecDeque;
use std::sync::{Arc, Mutex};

use betteroffice_xlsx::{
    MAX_COLLABORATION_BYTES, MAX_COLLABORATION_CLIENT_ID, UpdateEvent, UpdateOrigin,
    UpdateSubscription,
};
use wasm_bindgen::prelude::*;

use crate::core::Session;

/// days from the 1900 epoch to 1970-01-01, phantom-1900 leap day included.
const UNIX_EPOCH_SERIAL: f64 = 25569.0;
const MS_PER_DAY: f64 = 86_400_000.0;
const MAX_QUEUED_UPDATE_EVENTS: usize = 4_097;
const MAX_QUEUED_UPDATE_BYTES: usize = MAX_COLLABORATION_BYTES * 2 + MAX_QUEUED_UPDATE_EVENTS;

/// current utc time as a 1900-system serial for volatile cells; computed only
/// at the wasm boundary so the pure core stays deterministic and native-testable.
fn now_serial() -> Option<f64> {
    Some(js_sys::Date::now() / MS_PER_DAY + UNIX_EPOCH_SERIAL)
}

/// a workbook handle exposed to js; wraps the pure `Session`.
#[wasm_bindgen]
pub struct XlsxDocument {
    session: Session,
    update_observer: Option<UpdateObserver>,
}

struct UpdateObserver {
    pending: Arc<Mutex<PendingUpdateEvents>>,
    _subscription: UpdateSubscription,
}

#[derive(Default)]
struct PendingUpdateEvents {
    events: VecDeque<UpdateEvent>,
    bytes: usize,
    overflowed: bool,
}

#[wasm_bindgen]
impl XlsxDocument {
    /// open a workbook from raw `.xlsx` bytes.
    pub fn open(bytes: &[u8]) -> Result<XlsxDocument, JsValue> {
        Session::open(bytes, now_serial())
            .map(|session| XlsxDocument {
                session,
                update_observer: None,
            })
            .map_err(|e| JsValue::from_str(&e))
    }

    /// Open a replica with a positive, safe-integer client ID.
    #[wasm_bindgen(js_name = openCollaborative)]
    pub fn open_collaborative(bytes: &[u8], client_id: f64) -> Result<XlsxDocument, JsValue> {
        let client_id = parse_client_id(client_id)?;
        Session::open_collaborative(bytes, client_id, now_serial())
            .map(|session| XlsxDocument {
                session,
                update_observer: None,
            })
            .map_err(|e| JsValue::from_str(&e))
    }

    #[wasm_bindgen(getter, js_name = clientId)]
    pub fn client_id(&self) -> f64 {
        self.session.client_id() as f64
    }

    #[wasm_bindgen(js_name = encodeStateVector)]
    pub fn encode_state_vector(&self) -> Vec<u8> {
        self.session.encode_state_vector()
    }

    #[wasm_bindgen(js_name = encodeStateAsUpdate)]
    pub fn encode_state_as_update(&self) -> Vec<u8> {
        self.session.encode_state_as_update()
    }

    #[wasm_bindgen(js_name = encodeDiff)]
    pub fn encode_diff(&self, remote_state_vector: &[u8]) -> Result<Vec<u8>, JsValue> {
        self.session
            .encode_diff(remote_state_vector)
            .map_err(|e| JsValue::from_str(&e))
    }

    #[wasm_bindgen(js_name = applyUpdateJson)]
    pub fn apply_update_json(&mut self, update: &[u8]) -> Result<String, JsValue> {
        self.session
            .apply_update_json(update, now_serial())
            .map_err(|e| JsValue::from_str(&e))
    }

    /// Start queuing origin-prefixed Yrs update events for polling.
    #[wasm_bindgen(js_name = startUpdateObservation)]
    pub fn start_update_observation(&mut self) -> Result<(), JsValue> {
        if self.update_observer.is_some() {
            return Ok(());
        }
        let pending = Arc::new(Mutex::new(PendingUpdateEvents::default()));
        let observed = Arc::clone(&pending);
        let subscription = self
            .session
            .observe_update_v1(move |event| {
                let mut observed = observed
                    .lock()
                    .unwrap_or_else(|poisoned| poisoned.into_inner());
                if observed.overflowed {
                    return;
                }
                let bytes = event.update.len().saturating_add(1);
                if observed.events.len() >= MAX_QUEUED_UPDATE_EVENTS
                    || bytes > MAX_QUEUED_UPDATE_BYTES.saturating_sub(observed.bytes)
                {
                    observed.events.clear();
                    observed.bytes = 0;
                    observed.overflowed = true;
                    return;
                }
                observed.bytes += bytes;
                observed.events.push_back(event);
            })
            .map_err(|e| JsValue::from_str(&e))?;
        self.update_observer = Some(UpdateObserver {
            pending,
            _subscription: subscription,
        });
        Ok(())
    }

    /// Stop observation and discard queued events.
    #[wasm_bindgen(js_name = clearUpdateObservation)]
    pub fn clear_update_observation(&mut self) {
        self.update_observer = None;
    }

    /// Poll one event: origin byte (`0` local, `1` remote), then update; empty means none.
    #[wasm_bindgen(js_name = drainUpdateEvent)]
    pub fn drain_update_event(&self) -> Result<Vec<u8>, JsValue> {
        let Some(observer) = &self.update_observer else {
            return Ok(Vec::new());
        };
        let mut pending = observer
            .pending
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if pending.overflowed {
            pending.overflowed = false;
            return Err(JsValue::from_str(
                "xlsx update observation queue exceeded its limit",
            ));
        }
        let event = pending.events.pop_front();
        let Some(event) = event else {
            return Ok(Vec::new());
        };
        pending.bytes = pending
            .bytes
            .saturating_sub(event.update.len().saturating_add(1));
        let mut encoded = Vec::with_capacity(event.update.len() + 1);
        encoded.push(match event.origin {
            UpdateOrigin::Local => 0,
            UpdateOrigin::Remote => 1,
        });
        encoded.extend_from_slice(&event.update);
        Ok(encoded)
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

    #[wasm_bindgen(js_name = calculationStatusJson)]
    pub fn calculation_status_json(&self) -> Result<String, JsValue> {
        self.session
            .calculation_status_json()
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

    #[wasm_bindgen(js_name = cellPositionJson)]
    pub fn cell_position_json(&self, args: &str) -> Result<String, JsValue> {
        self.session
            .cell_position_json(args)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// a rectangular block of cells for clipboard copy.
    #[wasm_bindgen(js_name = rangeCellsJson)]
    pub fn range_cells_json(&self, args: &str) -> Result<String, JsValue> {
        self.session
            .range_cells_json(args)
            .map_err(|e| JsValue::from_str(&e))
    }

    #[wasm_bindgen(js_name = patchRangeStyleJson)]
    pub fn patch_range_style_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.session
            .patch_range_style_json(args, now_serial())
            .map_err(|e| JsValue::from_str(&e))
    }

    #[wasm_bindgen(js_name = setRangeNumberFormatJson)]
    pub fn set_range_number_format_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.session
            .set_range_number_format_json(args, now_serial())
            .map_err(|e| JsValue::from_str(&e))
    }

    #[wasm_bindgen(js_name = selectionFormattingJson)]
    pub fn selection_formatting_json(&self, args: &str) -> Result<String, JsValue> {
        self.session
            .selection_formatting_json(args)
            .map_err(|e| JsValue::from_str(&e))
    }

    #[wasm_bindgen(js_name = captureFormatJson)]
    pub fn capture_format_json(&self, args: &str) -> Result<String, JsValue> {
        self.session
            .capture_format_json(args)
            .map_err(|e| JsValue::from_str(&e))
    }

    #[wasm_bindgen(js_name = applyFormatJson)]
    pub fn apply_format_json(&mut self, args: &str) -> Result<String, JsValue> {
        self.session
            .apply_format_json(args, now_serial())
            .map_err(|e| JsValue::from_str(&e))
    }

    #[wasm_bindgen(js_name = mergedRangesJson)]
    pub fn merged_ranges_json(&self, args: &str) -> Result<String, JsValue> {
        self.session
            .merged_ranges_json(args)
            .map_err(|e| JsValue::from_str(&e))
    }

    #[wasm_bindgen(js_name = historyStateJson)]
    pub fn history_state_json(&self) -> Result<String, JsValue> {
        self.session
            .history_state_json()
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

fn parse_client_id(client_id: f64) -> Result<u64, JsValue> {
    if !client_id.is_finite()
        || client_id.fract() != 0.0
        || client_id < 1.0
        || client_id > MAX_COLLABORATION_CLIENT_ID as f64
    {
        return Err(JsValue::from_str(
            "client ID must be a nonzero integer no greater than Number.MAX_SAFE_INTEGER",
        ));
    }
    Ok(client_id as u64)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn update_queue_covers_the_largest_pending_resolution_burst() {
        assert_eq!(
            MAX_QUEUED_UPDATE_BYTES,
            MAX_COLLABORATION_BYTES * 2 + MAX_QUEUED_UPDATE_EVENTS
        );
    }
}
