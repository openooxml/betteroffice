use std::collections::VecDeque;
use std::sync::{Arc, Mutex};

use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;
use yrs::Subscription;

use crate::{
    CellSnapshot, DiagramSession, DiagramSnapshot, EditCtx, MAX_SAFE_CLIENT_ID, ShapeDraft,
    UpdateEvent, UpdateOrigin,
};
use vsdx_parse::{CellLocator, CellRow, CellSheet};

#[wasm_bindgen]
pub struct VsdxDocument {
    session: DiagramSession,
    update_observer: Option<UpdateObserver>,
}

struct UpdateObserver {
    pending: Arc<Mutex<PendingUpdates>>,
    _subscription: Subscription,
}

struct PendingUpdates {
    events: VecDeque<UpdateEvent>,
    resync_required: bool,
}

const MAX_PENDING_UPDATE_EVENTS: usize = 1024;

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct CellLocatorArgs {
    section: Option<String>,
    row_index: Option<u32>,
    row_name: Option<String>,
    cell_name: String,
}

impl TryFrom<CellLocatorArgs> for CellLocator {
    type Error = &'static str;

    fn try_from(value: CellLocatorArgs) -> Result<Self, Self::Error> {
        let row = match (value.row_index, value.row_name) {
            (Some(_), Some(_)) => {
                return Err("cell locator cannot contain both rowIndex and rowName");
            }
            (Some(index), None) => Some(CellRow::Index(index)),
            (None, Some(name)) => Some(CellRow::Name(name)),
            (None, None) => None,
        };
        Ok(Self {
            sheet: CellSheet::Page(0),
            shape_id: None,
            section: value.section,
            row,
            cell_name: value.cell_name,
        })
    }
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct SetCellFormulaArgs {
    page_id: String,
    shape_id: String,
    locator: CellLocatorArgs,
    formula: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct MoveShapeArgs {
    page_id: String,
    shape_id: String,
    x_formula: String,
    y_formula: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ResizeShapeArgs {
    page_id: String,
    shape_id: String,
    width_formula: String,
    height_formula: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ReorderShapeArgs {
    page_id: String,
    shape_id: String,
    to_index: u32,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ReorderPageArgs {
    page_id: String,
    to_index: u32,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct AddShapeArgs {
    page_id: String,
    draft: FormulaShapeDraft,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct FormulaShapeDraft {
    source_id: u32,
    name: Option<String>,
    cells: Vec<serde_json::Value>,
}

impl TryFrom<FormulaShapeDraft> for ShapeDraft {
    type Error = &'static str;

    fn try_from(value: FormulaShapeDraft) -> Result<Self, Self::Error> {
        let mut cells = Vec::with_capacity(value.cells.len());
        for cell in value.cells {
            if cell.get("value").is_some() {
                return Err("shape draft cells must not contain value");
            }
            cells.push(
                serde_json::from_value::<CellSnapshot>(cell)
                    .map_err(|_| "invalid shape draft cell")?,
            );
        }
        Ok(Self {
            source_id: value.source_id,
            name: value.name,
            cells,
        })
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct HistoryResult {
    applied: bool,
    snapshot: DiagramSnapshot,
}

#[wasm_bindgen]
impl VsdxDocument {
    #[wasm_bindgen(js_name = openCollaborative)]
    pub fn open_collaborative(bytes: &[u8], client_id: f64) -> Result<VsdxDocument, JsValue> {
        DiagramSession::open(bytes, parse_client_id(client_id)?)
            .map(|session| Self {
                session,
                update_observer: None,
            })
            .map_err(js_error)
    }

    #[wasm_bindgen(js_name = openCollaborativeFromUpdate)]
    pub fn open_collaborative_from_update(
        update: &[u8],
        client_id: f64,
    ) -> Result<VsdxDocument, JsValue> {
        DiagramSession::open_from_update(update, parse_client_id(client_id)?)
            .map(|session| Self {
                session,
                update_observer: None,
            })
            .map_err(js_error)
    }

    #[wasm_bindgen(getter, js_name = clientId)]
    pub fn client_id(&self) -> f64 {
        self.session.client_id() as f64
    }

    #[wasm_bindgen(js_name = snapshotJson)]
    pub fn snapshot_json(&self) -> Result<String, JsValue> {
        json(self.session.snapshot().map_err(js_error)?)
    }

    #[wasm_bindgen(js_name = mediaBytes)]
    pub fn media_bytes(&self, part_path: &str) -> Result<Vec<u8>, JsValue> {
        self.media_bytes_inner(part_path).map_err(js_error)
    }

    #[wasm_bindgen(js_name = encodeStateVector)]
    pub fn encode_state_vector(&self) -> Vec<u8> {
        self.session.encode_state_vector_v1()
    }

    #[wasm_bindgen(js_name = encodeStateAsUpdate)]
    pub fn encode_state_as_update(&self) -> Vec<u8> {
        self.session.encode_state_as_update_v1()
    }

    #[wasm_bindgen(js_name = encodeDiff)]
    pub fn encode_diff(&self, remote_state_vector: &[u8]) -> Result<Vec<u8>, JsValue> {
        self.session
            .encode_diff_v1(remote_state_vector)
            .map_err(js_error)
    }

    #[wasm_bindgen(js_name = applyUpdateJson)]
    pub fn apply_update_json(&self, update: &[u8]) -> Result<String, JsValue> {
        self.apply_update_json_inner(update).map_err(js_error)
    }

    #[wasm_bindgen(js_name = startUpdateObservation)]
    pub fn start_update_observation(&mut self) -> Result<(), JsValue> {
        if self.update_observer.is_some() {
            return Ok(());
        }
        let pending = Arc::new(Mutex::new(PendingUpdates {
            events: VecDeque::new(),
            resync_required: false,
        }));
        let observed = Arc::clone(&pending);
        let subscription = self
            .session
            .observe_update_v1(move |event| {
                let mut pending = observed
                    .lock()
                    .unwrap_or_else(|poisoned| poisoned.into_inner());
                if pending.events.len() == MAX_PENDING_UPDATE_EVENTS {
                    pending.events.clear();
                    pending.resync_required = true;
                }
                if !pending.resync_required {
                    pending.events.push_back(event);
                }
            })
            .map_err(js_error)?;
        self.update_observer = Some(UpdateObserver {
            pending,
            _subscription: subscription,
        });
        Ok(())
    }

    #[wasm_bindgen(js_name = clearUpdateObservation)]
    pub fn clear_update_observation(&mut self) {
        self.update_observer = None;
    }

    /// Returns `[2]` after overflow; discard queued observations and resync from a state vector.
    #[wasm_bindgen(js_name = drainUpdateEvent)]
    pub fn drain_update_event(&self) -> Vec<u8> {
        let Some(observer) = &self.update_observer else {
            return Vec::new();
        };
        let mut pending = observer
            .pending
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if pending.resync_required {
            pending.resync_required = false;
            return vec![2];
        }
        let Some(event) = pending.events.pop_front() else {
            return Vec::new();
        };
        let mut encoded = Vec::with_capacity(event.update.len() + 1);
        encoded.push(match event.origin {
            UpdateOrigin::Local => 0,
            UpdateOrigin::Remote => 1,
        });
        encoded.extend_from_slice(&event.update);
        encoded
    }

    #[wasm_bindgen(js_name = setCellFormulaJson)]
    pub fn set_cell_formula_json(&self, args: &str) -> Result<String, JsValue> {
        self.set_cell_formula_json_inner(args).map_err(js_error)
    }

    #[wasm_bindgen(js_name = moveShapeJson)]
    pub fn move_shape_json(&self, args: &str) -> Result<String, JsValue> {
        self.move_shape_json_inner(args).map_err(js_error)
    }

    #[wasm_bindgen(js_name = resizeShapeJson)]
    pub fn resize_shape_json(&self, args: &str) -> Result<String, JsValue> {
        self.resize_shape_json_inner(args).map_err(js_error)
    }

    #[wasm_bindgen(js_name = reorderShapeJson)]
    pub fn reorder_shape_json(&self, args: &str) -> Result<String, JsValue> {
        self.reorder_shape_json_inner(args).map_err(js_error)
    }

    #[wasm_bindgen(js_name = reorderPageJson)]
    pub fn reorder_page_json(&self, args: &str) -> Result<String, JsValue> {
        self.reorder_page_json_inner(args).map_err(js_error)
    }

    #[wasm_bindgen(js_name = addShapeJson)]
    pub fn add_shape_json(&self, args: &str) -> Result<String, JsValue> {
        self.add_shape_json_inner(args).map_err(js_error)
    }

    #[wasm_bindgen(js_name = undoJson)]
    pub fn undo_json(&self) -> Result<String, JsValue> {
        json(HistoryResult {
            applied: self.session.undo(),
            snapshot: self.session.snapshot().map_err(js_error)?,
        })
    }

    #[wasm_bindgen(js_name = redoJson)]
    pub fn redo_json(&self) -> Result<String, JsValue> {
        json(HistoryResult {
            applied: self.session.redo(),
            snapshot: self.session.snapshot().map_err(js_error)?,
        })
    }

    #[wasm_bindgen(js_name = canUndo)]
    pub fn can_undo(&self) -> bool {
        self.session.can_undo()
    }

    #[wasm_bindgen(js_name = canRedo)]
    pub fn can_redo(&self) -> bool {
        self.session.can_redo()
    }

    pub fn version() -> String {
        env!("CARGO_PKG_VERSION").to_owned()
    }
}

impl VsdxDocument {
    pub fn session(&self) -> &DiagramSession {
        &self.session
    }

    fn apply_update(&self, update: &[u8]) -> crate::EditResult<DiagramSnapshot> {
        self.session.apply_update_v1(update)
    }

    fn apply_update_json_inner(&self, update: &[u8]) -> Result<String, String> {
        self.apply_update(update)
            .map_err(|error| error.to_string())
            .and_then(json_inner)
    }

    fn media_bytes_inner(&self, part_path: &str) -> Result<Vec<u8>, String> {
        self.session
            .package()
            .map_err(|error| error.to_string())?
            .part_bytes(part_path)
            .map(ToOwned::to_owned)
            .ok_or_else(|| crate::EditError::InvalidState("media part was not found".to_owned()))
            .map_err(|error| error.to_string())
    }

    fn set_cell_formula(
        &self,
        args: SetCellFormulaArgs,
    ) -> crate::EditResult<crate::CellFormulaReceipt> {
        let locator = CellLocator::try_from(args.locator)
            .map_err(|error| crate::EditError::InvalidState(error.to_owned()))?;
        self.session.set_cell_formula_at(
            &local_context(),
            &args.page_id,
            &args.shape_id,
            locator,
            args.formula,
        )
    }

    fn set_cell_formula_json_inner(&self, args: &str) -> Result<String, String> {
        let args = parse_args_inner(args)?;
        self.set_cell_formula(args)
            .map_err(|error| error.to_string())
            .and_then(json_inner)
    }

    fn move_shape(&self, args: MoveShapeArgs) -> crate::EditResult<[crate::CellFormulaReceipt; 2]> {
        self.session.move_shape(
            &local_context(),
            &args.page_id,
            &args.shape_id,
            args.x_formula,
            args.y_formula,
        )
    }

    fn move_shape_json_inner(&self, args: &str) -> Result<String, String> {
        let args = parse_args_inner(args)?;
        self.move_shape(args)
            .map_err(|error| error.to_string())
            .and_then(json_inner)
    }

    fn reorder_shape_json_inner(&self, args: &str) -> Result<String, String> {
        let args: ReorderShapeArgs = parse_args_inner(args)?;
        self.session
            .reorder_shape(
                &local_context(),
                &args.page_id,
                &args.shape_id,
                args.to_index,
            )
            .map_err(|error| error.to_string())
            .and_then(json_inner)
    }

    fn reorder_page_json_inner(&self, args: &str) -> Result<String, String> {
        let args: ReorderPageArgs = parse_args_inner(args)?;
        self.session
            .reorder_page(&local_context(), &args.page_id, args.to_index)
            .map_err(|error| error.to_string())
            .and_then(json_inner)
    }

    fn add_shape_json_inner(&self, args: &str) -> Result<String, String> {
        let args: AddShapeArgs = parse_args_inner(args)?;
        let draft = args.draft.try_into().map_err(str::to_owned)?;
        self.session
            .add_shape(&local_context(), &args.page_id, &draft)
            .map_err(|error| error.to_string())
            .and_then(json_inner)
    }

    fn resize_shape(
        &self,
        args: ResizeShapeArgs,
    ) -> crate::EditResult<[crate::CellFormulaReceipt; 2]> {
        self.session.resize_shape(
            &local_context(),
            &args.page_id,
            &args.shape_id,
            args.width_formula,
            args.height_formula,
        )
    }

    fn resize_shape_json_inner(&self, args: &str) -> Result<String, String> {
        let args = parse_args_inner(args)?;
        self.resize_shape(args)
            .map_err(|error| error.to_string())
            .and_then(json_inner)
    }
}

fn local_context() -> EditCtx {
    EditCtx::local("wasm")
}

fn parse_args_inner<T: serde::de::DeserializeOwned>(args: &str) -> Result<T, String> {
    serde_json::from_str(args).map_err(|error| error.to_string())
}

fn json(value: impl Serialize) -> Result<String, JsValue> {
    serde_json::to_string(&value).map_err(js_error)
}

fn json_inner(value: impl Serialize) -> Result<String, String> {
    serde_json::to_string(&value).map_err(|error| error.to_string())
}

fn parse_client_id(client_id: f64) -> Result<u64, JsValue> {
    parse_client_id_raw(client_id).map_err(JsValue::from_str)
}

fn parse_client_id_raw(client_id: f64) -> Result<u64, &'static str> {
    if !client_id.is_finite()
        || client_id.fract() != 0.0
        || client_id < 1.0
        || client_id > MAX_SAFE_CLIENT_ID as f64
    {
        return Err("client ID must be a positive safe integer below Number.MAX_SAFE_INTEGER");
    }
    Ok(client_id as u64)
}

fn js_error(error: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&error.to_string())
}

#[cfg(test)]
mod tests {
    use super::{VsdxDocument, parse_client_id_raw};
    use crate::DiagramSession;
    use crate::diagram::MAX_SHAPE_NESTING;
    use crate::{MAX_SAFE_CLIENT_ID, PAGE_ORDER, PAGES, SHEETS};
    use yrs::{Array, ArrayPrelim, Map, MapPrelim, Out, ReadTxn, Transact};

    fn document() -> VsdxDocument {
        VsdxDocument::open_collaborative(
            include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx"),
            1.0,
        )
        .unwrap()
    }

    fn two_page_document() -> VsdxDocument {
        let seed = document();
        let session = DiagramSession::open_from_update(&seed.encode_state_as_update(), 2).unwrap();
        let mut txn = session.yrs_doc().transact_mut();
        let order = txn.get_array(PAGE_ORDER).unwrap();
        order.push_back(&mut txn, "page:2");
        let pages = txn.get_map(PAGES).unwrap();
        let page = pages.insert(&mut txn, "page:2", MapPrelim::default());
        page.insert(&mut txn, "id", "page:2");
        page.insert(&mut txn, "sourcePartPath", "visio/pages/page2.xml");
        page.insert(&mut txn, "shapes", ArrayPrelim::default());
        drop(txn);
        VsdxDocument {
            session,
            update_observer: None,
        }
    }

    fn add_cell(document: &VsdxDocument, key: &str, name: &str, formula: &str) {
        let mut txn = document.session().yrs_doc().transact_mut();
        let sheets = txn.get_map(SHEETS).unwrap();
        let shape = match sheets.get(&txn, "page:1:shape:1") {
            Some(Out::YMap(shape)) => shape,
            _ => unreachable!(),
        };
        let cells = match shape.get(&txn, "cells") {
            Some(Out::YMap(cells)) => cells,
            _ => unreachable!(),
        };
        let cell = cells.insert(&mut txn, key, MapPrelim::default());
        cell.insert(&mut txn, "name", name);
        cell.insert(&mut txn, "formula", formula);
        if name == "X" {
            cell.insert(&mut txn, "section", "Geometry");
            cell.insert(&mut txn, "rowIndex", 0.0);
        }
    }

    fn rewrite_formula_update(document: &VsdxDocument, key: &str, formula: &str) -> Vec<u8> {
        let attacker =
            DiagramSession::open_from_update(&document.encode_state_as_update(), 2).unwrap();
        let mut txn = attacker.yrs_doc().transact_mut();
        let sheets = txn.get_map(SHEETS).unwrap();
        let shape = match sheets.get(&txn, "page:1:shape:1") {
            Some(Out::YMap(shape)) => shape,
            _ => unreachable!(),
        };
        let cells = match shape.get(&txn, "cells") {
            Some(Out::YMap(cells)) => cells,
            _ => unreachable!(),
        };
        let cell = match cells.get(&txn, key) {
            Some(Out::YMap(cell)) => cell,
            _ => unreachable!(),
        };
        cell.insert(&mut txn, "formula", formula);
        drop(txn);
        attacker
            .encode_diff_v1(&document.encode_state_vector())
            .unwrap()
    }

    fn add_remote_cell_update(
        document: &VsdxDocument,
        formula: Option<&str>,
        value: Option<&str>,
    ) -> Vec<u8> {
        let attacker =
            DiagramSession::open_from_update(&document.encode_state_as_update(), 2).unwrap();
        let mut txn = attacker.yrs_doc().transact_mut();
        let sheets = txn.get_map(SHEETS).unwrap();
        let shape = match sheets.get(&txn, "page:1:shape:1") {
            Some(Out::YMap(shape)) => shape,
            _ => unreachable!(),
        };
        let cells = match shape.get(&txn, "cells") {
            Some(Out::YMap(cells)) => cells,
            _ => unreachable!(),
        };
        let cell = cells.insert(&mut txn, "Added", MapPrelim::default());
        cell.insert(&mut txn, "name", "Added");
        if let Some(formula) = formula {
            cell.insert(&mut txn, "formula", formula);
        }
        if let Some(value) = value {
            cell.insert(&mut txn, "value", value);
        }
        drop(txn);
        attacker
            .encode_diff_v1(&document.encode_state_vector())
            .unwrap()
    }

    #[test]
    fn client_ids_must_be_positive_safe_integers() {
        for client_id in [-1.0, 1.5, (MAX_SAFE_CLIENT_ID + 1) as f64] {
            assert!(parse_client_id_raw(client_id).is_err());
        }
    }

    #[test]
    fn set_cell_formula_json_inner_refuses_guarded_section_row_cell_edits() {
        let document = document();
        add_cell(&document, "Geometry\u{1f}IX:0\u{1f}X", "X", "GUARD(1)");
        assert_eq!(
            document.set_cell_formula_json_inner(r#"{"pageId":"page:1","shapeId":"page:1:shape:1","locator":{"section":"Geometry","rowIndex":0,"cellName":"X"},"formula":"2"}"#).unwrap_err(),
            "invalid diagram state: GUARD protects the requested cell"
        );
    }

    #[test]
    fn move_shape_json_inner_refuses_atomic_locks() {
        let document = document();
        add_cell(&document, "PinX", "PinX", "1");
        add_cell(&document, "PinY", "PinY", "1");
        add_cell(&document, "LockMoveY", "LockMoveY", "1");
        assert_eq!(
            document.move_shape_json_inner(
                r#"{"pageId":"page:1","shapeId":"page:1:shape:1","xFormula":"2","yFormula":"3"}"#
            ).unwrap_err(),
            "invalid diagram state: LockMoveY protects this move gesture"
        );
        let snapshot = document.snapshot_json().unwrap();
        assert!(snapshot.contains(r#""name":"PinX","formula":"1""#));
        assert!(snapshot.contains(r#""name":"PinY","formula":"1""#));

        add_cell(&document, "Width", "Width", "1");
        add_cell(&document, "Height", "Height", "1");
        add_cell(&document, "LockHeight", "LockHeight", "1");
        assert_eq!(
            document.resize_shape_json_inner(r#"{"pageId":"page:1","shapeId":"page:1:shape:1","widthFormula":"2","heightFormula":"3"}"#).unwrap_err(),
            "invalid diagram state: LockHeight protects this resize gesture"
        );
        let snapshot = document.snapshot_json().unwrap();
        assert!(snapshot.contains(r#""name":"Width","formula":"1""#));
        assert!(snapshot.contains(r#""name":"Height","formula":"1""#));
    }

    #[test]
    fn wasm_setatref_redirects_and_reports_the_target() {
        let document = document();
        add_cell(&document, "Width", "Width", "SETATREF(Target)");
        add_cell(&document, "Target", "Target", "1");
        assert_eq!(
            document
                .set_cell_formula_json(r#"{"pageId":"page:1","shapeId":"page:1:shape:1","locator":{"cellName":"Width"},"formula":"2"}"#)
                .unwrap(),
            r#"{"pageId":"page:1","shapeId":"page:1:shape:1","cellName":"Target","before":"1","after":"2"}"#
        );
        let snapshot = document.snapshot_json().unwrap();
        assert!(snapshot.contains(r#""name":"Width","formula":"SETATREF(Target)""#));
        assert!(snapshot.contains(r#""name":"Target","formula":"2""#));
    }

    #[test]
    fn wasm_move_shape_json_returns_receipts() {
        let document = document();
        add_cell(&document, "PinX", "PinX", "1");
        add_cell(&document, "PinY", "PinY", "1");
        assert_eq!(
            document
                .move_shape_json(
                    r#"{"pageId":"page:1","shapeId":"page:1:shape:1","xFormula":"2","yFormula":"3"}"#,
                )
                .unwrap(),
            r#"[{"pageId":"page:1","shapeId":"page:1:shape:1","cellName":"PinX","before":"1","after":"2"},{"pageId":"page:1","shapeId":"page:1:shape:1","cellName":"PinY","before":"1","after":"3"}]"#
        );
    }

    #[test]
    fn apply_update_json_inner_rejects_setatref_bypasses_without_changing_the_document() {
        let document = document();
        add_cell(&document, "Width", "Width", "SETATREF(Target)");
        add_cell(&document, "Target", "Target", "1");
        let before = document.encode_state_as_update();
        let attacker = DiagramSession::open_from_update(&before, 2).unwrap();
        let mut txn = attacker.yrs_doc().transact_mut();
        let sheets = txn.get_map(SHEETS).unwrap();
        let shape = match sheets.get(&txn, "page:1:shape:1") {
            Some(Out::YMap(shape)) => shape,
            _ => unreachable!(),
        };
        let cells = match shape.get(&txn, "cells") {
            Some(Out::YMap(cells)) => cells,
            _ => unreachable!(),
        };
        let width = match cells.get(&txn, "Width") {
            Some(Out::YMap(cell)) => cell,
            _ => unreachable!(),
        };
        width.insert(&mut txn, "formula", "2");
        drop(txn);
        let update = attacker
            .encode_diff_v1(&document.encode_state_vector())
            .unwrap();
        assert_eq!(
            document.apply_update_json_inner(&update).unwrap_err(),
            "invalid diagram state: remote update bypasses formula redirect at page:1/page:1:shape:1/Width"
        );
        assert_eq!(before, document.encode_state_as_update());
    }

    #[test]
    fn apply_update_json_inner_reports_decode_and_size_errors() {
        let document = document();
        assert_eq!(
            document.apply_update_json_inner(&[0]).unwrap_err(),
            "invalid yrs update: while trying to read more data (expected: 1 bytes), an unexpected end of buffer was reached"
        );
        assert_eq!(
            document
                .apply_update_json_inner(&vec![0; crate::MAX_UPDATE_BYTES + 1])
                .unwrap_err(),
            "invalid yrs update: update exceeds 67108864 bytes"
        );
    }

    #[test]
    fn set_cell_formula_json_inner_reports_unevaluable_lock_and_malformed_formula() {
        let locked_document = document();
        add_cell(&locked_document, "PinX", "PinX", "1");
        add_cell(&locked_document, "PinY", "PinY", "1");
        add_cell(&locked_document, "LockMoveX", "LockMoveX", "Unknown()");
        assert_eq!(
            locked_document
                .move_shape_json_inner(
                    r#"{"pageId":"page:1","shapeId":"page:1:shape:1","xFormula":"2","yFormula":"3"}"#
                )
                .unwrap_err(),
            "invalid diagram state: cannot evaluate LockMoveX"
        );
        let document = document();
        add_cell(&document, "PinX", "PinX", "1+");
        assert_eq!(
            document
                .set_cell_formula_json_inner(
                    r#"{"pageId":"page:1","shapeId":"page:1:shape:1","locator":{"cellName":"PinX"},"formula":"2"}"#
                )
                .unwrap_err(),
            "invalid diagram state: cannot inspect existing formula: expected expression"
        );
    }

    #[test]
    fn apply_update_json_inner_refuses_protected_and_malformed_formula_changes() {
        let guarded = document();
        add_cell(&guarded, "Width", "Width", "GUARD(1)");
        assert_eq!(
            guarded
                .apply_update_json_inner(&rewrite_formula_update(&guarded, "Width", "2"))
                .unwrap_err(),
            "invalid diagram state: GUARD protects the requested cell"
        );

        let locked = document();
        add_cell(&locked, "PinX", "PinX", "1");
        add_cell(&locked, "LockMoveX", "LockMoveX", "1");
        assert_eq!(
            locked
                .apply_update_json_inner(&rewrite_formula_update(&locked, "PinX", "2"))
                .unwrap_err(),
            "invalid diagram state: LockMoveX protects this move gesture"
        );

        let unevaluable = document();
        add_cell(&unevaluable, "PinX", "PinX", "1");
        add_cell(&unevaluable, "LockMoveX", "LockMoveX", "Unknown()");
        assert_eq!(
            unevaluable
                .apply_update_json_inner(&rewrite_formula_update(&unevaluable, "PinX", "2"))
                .unwrap_err(),
            "invalid diagram state: cannot evaluate LockMoveX"
        );

        let malformed = document();
        add_cell(&malformed, "Width", "Width", "1+");
        assert_eq!(
            malformed
                .apply_update_json_inner(&rewrite_formula_update(&malformed, "Width", "2"))
                .unwrap_err(),
            "invalid diagram state: cannot inspect existing formula: expected expression"
        );
    }

    #[test]
    fn apply_update_json_inner_refuses_frozen_metadata_changes() {
        let document = document();
        let attacker =
            DiagramSession::open_from_update(&document.encode_state_as_update(), 2).unwrap();
        let mut txn = attacker.yrs_doc().transact_mut();
        txn.get_map(crate::META)
            .unwrap()
            .insert(&mut txn, "fingerprint", "changed");
        drop(txn);
        let update = attacker
            .encode_diff_v1(&document.encode_state_vector())
            .unwrap();
        assert_eq!(
            document.apply_update_json_inner(&update).unwrap_err(),
            "invalid diagram state: remote update changes immutable diagram metadata fingerprint"
        );
    }

    #[test]
    fn wasm_media_bytes_returns_committed_media() {
        let document = VsdxDocument::open_collaborative(
            include_bytes!("../../vsdx-parse/tests/fixtures/nested-groups.vsdx"),
            1.0,
        )
        .unwrap();
        assert_eq!(
            document
                .media_bytes_inner("visio/media/image1.png")
                .unwrap(),
            vec![137, 80, 78, 71, 13, 10, 26, 10]
        );
    }

    #[test]
    fn wasm_add_shape_json_returns_a_receipt() {
        let document = document();
        assert_eq!(
            document
                .add_shape_json(
                    r#"{"pageId":"page:1","draft":{"sourceId":2,"name":"Added","cells":[]}}"#
                )
                .unwrap(),
            r#"{"pageId":"page:1","shapeId":"page:1:shape:1:0","fromIndex":null,"toIndex":1}"#
        );
    }

    #[test]
    fn media_bytes_inner_rejects_unknown_parts() {
        assert_eq!(
            document()
                .media_bytes_inner("visio/media/missing.png")
                .unwrap_err(),
            "invalid diagram state: media part was not found"
        );
    }

    #[test]
    fn add_shape_json_inner_rejects_raw_values() {
        assert_eq!(
            document()
                .add_shape_json_inner(
                    r#"{"pageId":"page:1","draft":{"sourceId":3,"cells":[{"value":"1"}]}}"#
                )
                .unwrap_err(),
            "shape draft cells must not contain value"
        );
    }

    #[test]
    fn remote_cell_additions_must_be_formula_only() {
        let raw = document();
        assert_eq!(
            raw.apply_update_json_inner(&add_remote_cell_update(&raw, None, Some("1")))
                .unwrap_err(),
            "invalid diagram state: remote update adds untrusted cached cell value"
        );
        let formula = document();
        assert!(
            formula
                .apply_update_json_inner(&add_remote_cell_update(&formula, Some("1"), None))
                .is_ok()
        );
    }

    #[test]
    fn remote_updates_reject_shape_nesting_beyond_the_limit() {
        let document = document();
        let attacker =
            DiagramSession::open_from_update(&document.encode_state_as_update(), 2).unwrap();
        let mut txn = attacker.yrs_doc().transact_mut();
        let sheets = txn.get_map(SHEETS).unwrap();
        let root = match sheets.get(&txn, "page:1:shape:1") {
            Some(Out::YMap(shape)) => shape,
            _ => unreachable!(),
        };
        let mut parent_id = "page:1:shape:1".to_owned();
        for index in 0..MAX_SHAPE_NESTING {
            let shape_id = format!("page:1:shape:deep:{index}");
            let shape = sheets.insert(&mut txn, shape_id.as_str(), MapPrelim::default());
            shape.insert(&mut txn, "id", shape_id.as_str());
            shape.insert(&mut txn, "pageId", "page:1");
            shape.insert(&mut txn, "sourceId", (index + 10) as f64);
            shape.insert(&mut txn, "parentId", parent_id.as_str());
            shape.insert(&mut txn, "cells", MapPrelim::default());
            shape.insert(&mut txn, "shapes", ArrayPrelim::default());
            if index == 0 {
                let roots = match root.get(&txn, "shapes") {
                    Some(Out::YArray(shapes)) => shapes,
                    _ => unreachable!(),
                };
                roots.push_back(&mut txn, shape_id.as_str());
            } else {
                let parent = match sheets.get(&txn, &parent_id) {
                    Some(Out::YMap(shape)) => shape,
                    _ => unreachable!(),
                };
                let parent_children = match parent.get(&txn, "shapes") {
                    Some(Out::YArray(shapes)) => shapes,
                    _ => unreachable!(),
                };
                parent_children.push_back(&mut txn, shape_id.as_str());
            }
            parent_id = shape_id;
        }
        drop(txn);
        let update = attacker
            .encode_diff_v1(&document.encode_state_vector())
            .unwrap();
        assert_eq!(
            document.apply_update_json_inner(&update).unwrap_err(),
            "invalid diagram state: shape nesting exceeds maximum depth"
        );
    }

    #[test]
    fn remote_updates_reject_shape_attached_to_two_pages() {
        let document = two_page_document();
        let attacker =
            DiagramSession::open_from_update(&document.encode_state_as_update(), 3).unwrap();
        let mut txn = attacker.yrs_doc().transact_mut();
        let sheets = txn.get_map(SHEETS).unwrap();
        let root = match sheets.get(&txn, "page:1:shape:1") {
            Some(Out::YMap(shape)) => shape,
            _ => unreachable!(),
        };
        root.remove(&mut txn, "pageId");
        let pages = txn.get_map(PAGES).unwrap();
        let page = match pages.get(&txn, "page:2") {
            Some(Out::YMap(page)) => page,
            _ => unreachable!(),
        };
        let shapes = match page.get(&txn, "shapes") {
            Some(Out::YArray(shapes)) => shapes,
            _ => unreachable!(),
        };
        shapes.push_back(&mut txn, "page:1:shape:1");
        drop(txn);
        let update = attacker
            .encode_diff_v1(&document.encode_state_vector())
            .unwrap();
        assert_eq!(
            document.apply_update_json_inner(&update).unwrap_err(),
            "invalid diagram state: shape is attached to multiple pages"
        );
    }

    #[test]
    fn remote_updates_reject_non_string_parent_id() {
        let document = document();
        let attacker =
            DiagramSession::open_from_update(&document.encode_state_as_update(), 2).unwrap();
        let mut txn = attacker.yrs_doc().transact_mut();
        let sheets = txn.get_map(SHEETS).unwrap();
        let shape = sheets.insert(
            &mut txn,
            "page:1:shape:invalid-parent",
            MapPrelim::default(),
        );
        shape.insert(&mut txn, "id", "page:1:shape:invalid-parent");
        shape.insert(&mut txn, "pageId", "page:1");
        shape.insert(&mut txn, "sourceId", 99.0);
        shape.insert(&mut txn, "parentId", 1.0);
        shape.insert(&mut txn, "cells", MapPrelim::default());
        shape.insert(&mut txn, "shapes", ArrayPrelim::default());
        let pages = txn.get_map(PAGES).unwrap();
        let page = match pages.get(&txn, "page:1") {
            Some(Out::YMap(page)) => page,
            _ => unreachable!(),
        };
        let roots = match page.get(&txn, "shapes") {
            Some(Out::YArray(shapes)) => shapes,
            _ => unreachable!(),
        };
        roots.push_back(&mut txn, "page:1:shape:invalid-parent");
        drop(txn);
        let update = attacker
            .encode_diff_v1(&document.encode_state_vector())
            .unwrap();
        assert_eq!(
            document.apply_update_json_inner(&update).unwrap_err(),
            "invalid diagram state: shape parentId is not a string"
        );
    }

    #[test]
    fn remote_updates_cannot_detach_grouped_children() {
        let document = VsdxDocument::open_collaborative(
            include_bytes!("../../vsdx-parse/tests/fixtures/nested-groups.vsdx"),
            1.0,
        )
        .unwrap();
        let child_id = document.session().snapshot().unwrap().pages[0].shapes[0].children[0]
            .id
            .clone();
        let attacker =
            DiagramSession::open_from_update(&document.encode_state_as_update(), 2).unwrap();
        let mut txn = attacker.yrs_doc().transact_mut();
        let sheets = txn.get_map(SHEETS).unwrap();
        let child = match sheets.get(&txn, &child_id) {
            Some(Out::YMap(child)) => child,
            _ => unreachable!(),
        };
        child.insert(&mut txn, "parentId", "page:1:shape:detached");
        drop(txn);
        let update = attacker
            .encode_diff_v1(&document.encode_state_vector())
            .unwrap();
        assert_eq!(
            document.apply_update_json_inner(&update).unwrap_err(),
            "invalid diagram state: shape parentId does not match shape order"
        );
    }

    #[test]
    fn wasm_update_observation_overflow_requires_resync() {
        let mut document = document();
        document.start_update_observation().unwrap();
        for index in 0..=super::MAX_PENDING_UPDATE_EVENTS {
            add_cell(
                &document,
                &format!("Overflow{index}"),
                &format!("Overflow{index}"),
                "1",
            );
        }
        assert_eq!(document.drain_update_event(), vec![2]);
    }
}
