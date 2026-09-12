//! Collaborative yrs-backed VSDX diagram model.

use std::cell::RefCell;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::sync::atomic::{AtomicU64, Ordering};

use sha2::{Digest, Sha256};
use yrs::updates::decoder::{Decode, Decoder, DecoderV1};
use yrs::updates::encoder::Encode;
use yrs::{
    ClientID, Doc, OffsetKind, Options, ReadTxn, StateVector, Subscription, Transact, Update,
    WriteTxn,
};

mod diagram;
mod model;
mod undo;

pub use model::*;
pub use undo::DiagramUndoManager;

pub(crate) const META: &str = "vsdx:meta";
pub(crate) const PAGE_ORDER: &str = "vsdx:page-order";
pub(crate) const PAGES: &str = "vsdx:pages";
pub(crate) const SHEETS: &str = "vsdx:sheets";
pub(crate) const STORIES: &str = "vsdx:stories";
pub(crate) const REMOTE_ORIGIN: &str = "vsdx:remote";
pub(crate) const HYDRATE_ORIGIN: &str = "vsdx:hydrate";
const BOOTSTRAP_CLIENT_ID: u64 = (1_u64 << 53) - 1;
pub const MAX_SAFE_CLIENT_ID: u64 = BOOTSTRAP_CLIENT_ID - 1;
pub const MAX_UPDATE_BYTES: usize = 64 * 1024 * 1024;
const MAX_STATE_VECTOR_ENTRIES: u32 = 65_536;
const MAX_STATE_VECTOR_BYTES: usize = 1024 * 1024;

pub struct DiagramSession {
    pub(crate) doc: Doc,
    client_id: u64,
    id_counter: AtomicU64,
    undo: RefCell<DiagramUndoManager>,
}

impl DiagramSession {
    pub fn open(bytes: &[u8], client_id: u64) -> EditResult<Self> {
        let package =
            vsdx_parse::parse_vsdx(bytes).map_err(|error| EditError::Parse(error.to_string()))?;
        Self::from_package_with_fingerprint(
            package,
            format!("{:x}", Sha256::digest(bytes)),
            client_id,
        )
    }

    pub fn from_package(package: vsdx_parse::VsdxPackage, client_id: u64) -> EditResult<Self> {
        let json =
            serde_json::to_vec(&package).map_err(|error| EditError::Json(error.to_string()))?;
        Self::from_package_with_fingerprint(
            package,
            format!("{:x}", Sha256::digest(json)),
            client_id,
        )
    }

    fn from_package_with_fingerprint(
        package: vsdx_parse::VsdxPackage,
        fingerprint: String,
        client_id: u64,
    ) -> EditResult<Self> {
        validate_client_id(client_id)?;
        let bootstrap = doc_with_client_id(BOOTSTRAP_CLIENT_ID);
        diagram::seed_doc(&bootstrap, &package, &fingerprint)?;
        let baseline = bootstrap
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let doc = doc_with_client_id(client_id);
        hydrate_doc(&doc, &baseline)?;
        diagram::validate_doc(&doc)?;
        let undo = DiagramUndoManager::new(&doc, client_id)?;
        let id_counter = diagram::next_id_counter(&doc, client_id);
        Ok(Self {
            doc,
            client_id,
            id_counter: AtomicU64::new(id_counter),
            undo: RefCell::new(undo),
        })
    }

    pub fn open_from_update(update: &[u8], client_id: u64) -> EditResult<Self> {
        validate_client_id(client_id)?;
        if update.len() > MAX_UPDATE_BYTES {
            return Err(EditError::InvalidUpdate(format!(
                "update exceeds {MAX_UPDATE_BYTES} bytes"
            )));
        }
        let doc = doc_with_client_id(client_id);
        hydrate_doc(&doc, update)?;
        diagram::validate_doc(&doc)?;
        let undo = DiagramUndoManager::new(&doc, client_id)?;
        let id_counter = diagram::next_id_counter(&doc, client_id);
        Ok(Self {
            doc,
            client_id,
            id_counter: AtomicU64::new(id_counter),
            undo: RefCell::new(undo),
        })
    }

    pub fn client_id(&self) -> u64 {
        self.client_id
    }
    pub fn yrs_doc(&self) -> &Doc {
        &self.doc
    }
    pub fn encode_state_vector_v1(&self) -> Vec<u8> {
        self.doc.transact().state_vector().encode_v1()
    }
    pub fn encode_state_as_update_v1(&self) -> Vec<u8> {
        self.doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default())
    }
    pub fn encode_diff_v1(&self, remote: &[u8]) -> EditResult<Vec<u8>> {
        let vector = decode_state_vector_v1(remote).map_err(EditError::InvalidStateVector)?;
        Ok(self.doc.transact().encode_diff_v1(&vector))
    }

    pub fn apply_update_v1(&self, bytes: &[u8]) -> EditResult<DiagramSnapshot> {
        if bytes.len() > MAX_UPDATE_BYTES {
            return Err(EditError::InvalidUpdate(format!(
                "update exceeds {MAX_UPDATE_BYTES} bytes"
            )));
        }
        let incoming = decode_update_v1(bytes).map_err(EditError::InvalidUpdate)?;
        let staged = doc_with_client_id(self.client_id);
        hydrate_doc(&staged, &self.encode_state_as_update_v1())?;
        staged
            .transact_mut_with(REMOTE_ORIGIN)
            .apply_update(incoming)
            .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
        // Remote bytes are trusted only after decoding and schema/policy validation on a staged clone;
        // this rejects malformed state and protected-cell rewrites, but cannot authenticate an author.
        diagram::validate_remote_update(&self.doc, &staged)?;
        self.doc
            .transact_mut_with(REMOTE_ORIGIN)
            .apply_update(decode_update_v1(bytes).map_err(EditError::InvalidUpdate)?)
            .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
        self.snapshot()
    }

    pub fn observe_update_v1<F>(&self, callback: F) -> EditResult<Subscription>
    where
        F: Fn(UpdateEvent) + 'static,
    {
        self.doc
            .observe_update_v1(move |txn, event| {
                let origin = if txn
                    .origin()
                    .is_some_and(|origin| origin.as_ref() == REMOTE_ORIGIN.as_bytes())
                {
                    UpdateOrigin::Remote
                } else {
                    UpdateOrigin::Local
                };
                let _ = catch_unwind(AssertUnwindSafe(|| {
                    callback(UpdateEvent {
                        update: event.update.clone(),
                        origin,
                    })
                }));
            })
            .map_err(|error| EditError::Observer(error.to_string()))
    }

    pub fn undo(&self) -> bool {
        self.undo.borrow_mut().undo()
    }
    pub fn redo(&self) -> bool {
        self.undo.borrow_mut().redo()
    }
    pub fn can_undo(&self) -> bool {
        self.undo.borrow().can_undo()
    }
    pub fn can_redo(&self) -> bool {
        self.undo.borrow().can_redo()
    }
    pub fn add_undo_barrier(&self) {
        self.undo.borrow_mut().add_undo_barrier()
    }
    pub(crate) fn transact_for(&self, context: &EditCtx) -> yrs::TransactionMut<'_> {
        match context.origin {
            EditOrigin::Local => self.doc.transact_mut_with(self.client_id),
            EditOrigin::Agent => self.doc.transact_mut_with("vsdx:agent"),
            EditOrigin::Remote => self.doc.transact_mut_with(REMOTE_ORIGIN),
            EditOrigin::System => self.doc.transact_mut_with("vsdx:system"),
        }
    }
    pub(crate) fn next_id(&self, prefix: &str) -> String {
        format!(
            "{prefix}:{}:{}",
            self.client_id,
            self.id_counter.fetch_add(1, Ordering::Relaxed)
        )
    }
}

fn doc_with_client_id(client_id: u64) -> Doc {
    let mut options = Options::with_client_id(ClientID::new(client_id));
    options.offset_kind = OffsetKind::Utf16;
    Doc::with_options(options)
}
fn validate_client_id(client_id: u64) -> EditResult<()> {
    if client_id == 0 || client_id > MAX_SAFE_CLIENT_ID {
        Err(EditError::InvalidClientId(client_id))
    } else {
        Ok(())
    }
}
fn hydrate_doc(doc: &Doc, bytes: &[u8]) -> EditResult<()> {
    let update = decode_update_v1(bytes).map_err(EditError::InvalidUpdate)?;
    let mut txn = doc.transact_mut_with(HYDRATE_ORIGIN);
    txn.apply_update(update)
        .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
    txn.get_or_insert_array(PAGE_ORDER);
    for root in [META, PAGES, SHEETS, STORIES] {
        txn.get_or_insert_map(root);
    }
    Ok(())
}
fn decode_update_v1(bytes: &[u8]) -> Result<Update, String> {
    let mut decoder = DecoderV1::from(bytes);
    let update = Update::decode(&mut decoder).map_err(|error| error.to_string())?;
    if !decoder
        .read_to_end()
        .map_err(|error| error.to_string())?
        .is_empty()
    {
        return Err("update contains trailing bytes".to_owned());
    }
    Ok(update)
}
fn decode_state_vector_v1(bytes: &[u8]) -> Result<StateVector, String> {
    if bytes.len() > MAX_STATE_VECTOR_BYTES {
        return Err(format!(
            "state vector exceeds {MAX_STATE_VECTOR_BYTES} bytes"
        ));
    }
    validate_state_vector_entry_count(bytes)?;
    let mut decoder = DecoderV1::from(bytes);
    let vector = StateVector::decode(&mut decoder).map_err(|error| error.to_string())?;
    if !decoder
        .read_to_end()
        .map_err(|error| error.to_string())?
        .is_empty()
    {
        return Err("state vector contains trailing bytes".to_owned());
    }
    Ok(vector)
}
fn validate_state_vector_entry_count(bytes: &[u8]) -> Result<(), String> {
    let Some((&first, _)) = bytes.split_first() else {
        return Err("state vector is empty".to_owned());
    };
    let mut value = u32::from(first & 0x7f);
    let mut shift = 7;
    let mut used = 1;
    let mut byte = first;
    while byte & 0x80 != 0 {
        if used == 5 || used >= bytes.len() {
            return Err("invalid state vector entry count".to_owned());
        }
        byte = bytes[used];
        if used == 4 && byte > 0x0f {
            return Err("invalid state vector entry count".to_owned());
        }
        value |= u32::from(byte & 0x7f) << shift;
        shift += 7;
        used += 1;
    }
    if value > MAX_STATE_VECTOR_ENTRIES {
        return Err(format!(
            "state vector contains {value} entries, exceeds the {MAX_STATE_VECTOR_ENTRIES}-entry limit"
        ));
    }
    if value as usize > bytes.len().saturating_sub(used) / 2 {
        return Err("state vector entry count exceeds its payload".to_owned());
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use vsdx_parse::{CellLocator, CellRow, CellSheet};
    use yrs::{Any, Array, ArrayPrelim, Map, MapPrelim, Transact};

    fn session() -> DiagramSession {
        let doc = doc_with_client_id(7);
        let mut txn = doc.transact_mut_with(HYDRATE_ORIGIN);
        let meta = txn.get_or_insert_map(META);
        meta.insert(&mut txn, "schemaVersion", 1.0);
        meta.insert(&mut txn, "fingerprint", "test");
        meta.insert(
            &mut txn,
            "packageJson",
            Any::Buffer(std::sync::Arc::from([])),
        );
        let pages = txn.get_or_insert_map(PAGES);
        let sheets = txn.get_or_insert_map(SHEETS);
        let order = txn.get_or_insert_array(PAGE_ORDER);
        txn.get_or_insert_map(STORIES);
        for page_id in ["page:1", "page:2"] {
            order.push_back(&mut txn, page_id);
            let page = pages.insert(&mut txn, page_id, MapPrelim::default());
            page.insert(&mut txn, "id", page_id);
            page.insert(&mut txn, "sourcePartPath", format!("/{page_id}"));
            page.insert(&mut txn, "shapes", ArrayPrelim::default());
        }
        for (id, page_id) in [("page:1:shape:1", "page:1"), ("page:1:shape:2", "page:1")] {
            let shape = sheets.insert(&mut txn, id, MapPrelim::default());
            shape.insert(&mut txn, "id", id);
            shape.insert(&mut txn, "pageId", page_id);
            shape.insert(&mut txn, "origin", "package");
            shape.insert(&mut txn, "sourceId", 1.0);
            shape.insert(&mut txn, "cells", MapPrelim::default());
            let page = match pages.get(&txn, page_id) {
                Some(yrs::Out::YMap(page)) => page,
                _ => unreachable!(),
            };
            let shapes = match page.get(&txn, "shapes") {
                Some(yrs::Out::YArray(shapes)) => shapes,
                _ => unreachable!(),
            };
            shapes.push_back(&mut txn, id);
        }
        drop(txn);
        DiagramSession {
            undo: std::cell::RefCell::new(DiagramUndoManager::new(&doc, 7).unwrap()),
            doc,
            client_id: 7,
            id_counter: AtomicU64::new(0),
        }
    }

    fn add_cell(session: &DiagramSession, name: &str, formula: Option<&str>, value: Option<&str>) {
        add_cell_at(session, name, None, None, formula, value);
    }

    fn add_cell_at(
        session: &DiagramSession,
        name: &str,
        section: Option<&str>,
        row: Option<CellRow>,
        formula: Option<&str>,
        value: Option<&str>,
    ) {
        add_shape_cell(
            session,
            "page:1:shape:1",
            name,
            section,
            row,
            formula,
            value,
        );
    }

    fn add_child_shape(session: &DiagramSession, id: &str, parent_id: &str) {
        let mut txn = session.doc.transact_mut_with(HYDRATE_ORIGIN);
        let sheets = txn.get_map(SHEETS).unwrap();
        let shape = sheets.insert(&mut txn, id, MapPrelim::default());
        shape.insert(&mut txn, "id", id);
        shape.insert(&mut txn, "pageId", "page:1");
        shape.insert(&mut txn, "origin", "package");
        shape.insert(&mut txn, "sourceId", 3.0);
        shape.insert(&mut txn, "parentId", parent_id);
        shape.insert(&mut txn, "cells", MapPrelim::default());
    }

    fn shape_cells<T: yrs::ReadTxn>(txn: &T, shape_id: &str) -> yrs::MapRef {
        let sheets = txn.get_map(SHEETS).unwrap();
        let shape = match sheets.get(txn, shape_id) {
            Some(yrs::Out::YMap(shape)) => shape,
            _ => unreachable!(),
        };
        match shape.get(txn, "cells") {
            Some(yrs::Out::YMap(cells)) => cells,
            _ => unreachable!(),
        }
    }

    fn peer_doc(session: &DiagramSession, client_id: u64) -> Doc {
        let doc = doc_with_client_id(client_id);
        hydrate_doc(&doc, &session.encode_state_as_update_v1()).unwrap();
        doc
    }

    fn peer_update(session: &DiagramSession, peer: &Doc) -> Vec<u8> {
        peer.transact()
            .encode_diff_v1(&session.doc.transact().state_vector())
    }

    fn write_peer_cell_field(peer: &Doc, shape_id: &str, cell: &str, field: &str, value: &str) {
        let mut txn = peer.transact_mut();
        let cells = shape_cells(&txn, shape_id);
        let cell = match cells.get(&txn, cell) {
            Some(yrs::Out::YMap(cell)) => cell,
            _ => unreachable!(),
        };
        cell.insert(&mut txn, field, value);
    }

    fn write_peer_shape_field(peer: &Doc, shape_id: &str, field: &str, value: &str) {
        let mut txn = peer.transact_mut();
        let sheets = txn.get_map(SHEETS).unwrap();
        let shape = match sheets.get(&txn, shape_id) {
            Some(yrs::Out::YMap(shape)) => shape,
            _ => unreachable!(),
        };
        shape.insert(&mut txn, field, value);
    }

    fn write_peer_new_cell(peer: &Doc, shape_id: &str, name: &str, formula: &str) {
        let mut txn = peer.transact_mut();
        let cells = shape_cells(&txn, shape_id);
        let cell = cells.insert(&mut txn, name, MapPrelim::default());
        cell.insert(&mut txn, "name", name);
        cell.insert(&mut txn, "formula", formula);
    }

    fn write_peer_new_shape(
        peer: &Doc,
        shape_id: &str,
        page_id: &str,
        origin: &str,
        source_id: f64,
    ) {
        let mut txn = peer.transact_mut();
        let sheets = txn.get_map(SHEETS).unwrap();
        let shape = sheets.insert(&mut txn, shape_id, MapPrelim::default());
        shape.insert(&mut txn, "id", shape_id);
        shape.insert(&mut txn, "pageId", page_id);
        shape.insert(&mut txn, "origin", origin);
        shape.insert(&mut txn, "sourceId", source_id);
        shape.insert(&mut txn, "cells", MapPrelim::default());
        let pages = txn.get_map(PAGES).unwrap();
        let page = match pages.get(&txn, page_id) {
            Some(yrs::Out::YMap(page)) => page,
            _ => unreachable!(),
        };
        let shapes = match page.get(&txn, "shapes") {
            Some(yrs::Out::YArray(shapes)) => shapes,
            _ => unreachable!(),
        };
        shapes.push_back(&mut txn, shape_id);
    }

    fn add_shape_cell(
        session: &DiagramSession,
        shape_id: &str,
        name: &str,
        section: Option<&str>,
        row: Option<CellRow>,
        formula: Option<&str>,
        value: Option<&str>,
    ) {
        let mut txn = session.doc.transact_mut_with(HYDRATE_ORIGIN);
        let cells = shape_cells(&txn, shape_id);
        let key = match (&section, &row) {
            (Some(section), Some(CellRow::Index(row))) => {
                format!("{section}\u{1f}IX:{row}\u{1f}{name}")
            }
            (Some(section), Some(CellRow::Name(row))) => {
                format!("{section}\u{1f}N:{row}\u{1f}{name}")
            }
            _ => name.to_owned(),
        };
        let cell = cells.insert(&mut txn, key.as_str(), MapPrelim::default());
        cell.insert(&mut txn, "name", name);
        if let Some(section) = section {
            cell.insert(&mut txn, "section", section);
        }
        if let Some(row) = row {
            match row {
                CellRow::Index(row) => cell.insert(&mut txn, "rowIndex", row as f64),
                CellRow::Name(row) => cell.insert(&mut txn, "rowName", row),
            };
        }
        if let Some(formula) = formula {
            cell.insert(&mut txn, "formula", formula);
            cell.insert(&mut txn, "baselineFormula", formula);
        }
        if let Some(value) = value {
            cell.insert(&mut txn, "value", value);
        }
    }

    #[test]
    fn guards_refuse_all_formula_spellings() {
        for formula in ["GUARD(1)", "=GUARD(1)", "guard(1)", "IF(1, GUARD(1), 0)"] {
            let session = session();
            add_cell(&session, "Width", Some(formula), None);
            assert!(
                session
                    .set_cell_formula(
                        &EditCtx::local("a"),
                        "page:1",
                        "page:1:shape:1",
                        "Width",
                        "2"
                    )
                    .is_err()
            );
        }
    }

    #[test]
    fn matching_locks_refuse_move_and_resize() {
        let session = session();
        add_cell(&session, "PinX", Some("1"), None);
        add_cell(&session, "PinY", Some("1"), None);
        add_cell(&session, "LockMoveX", Some("1"), None);
        assert!(
            session
                .move_shape(&EditCtx::local("a"), "page:1", "page:1:shape:1", "2", "3")
                .is_err()
        );
        add_cell(&session, "Width", Some("1"), None);
        add_cell(&session, "Height", Some("1"), None);
        add_cell(&session, "LockWidth", Some("1"), None);
        assert!(
            session
                .resize_shape(&EditCtx::local("a"), "page:1", "page:1:shape:1", "2", "3")
                .is_err()
        );
    }

    #[test]
    fn move_refusal_leaves_both_axes_unchanged() {
        let session = session();
        add_cell(&session, "PinX", Some("1"), None);
        add_cell(&session, "PinY", Some("1"), None);
        add_cell(&session, "LockMoveY", Some("1"), None);
        assert!(
            session
                .move_shape(&EditCtx::local("a"), "page:1", "page:1:shape:1", "2", "3")
                .is_err()
        );
        let cells = &session.snapshot().unwrap().pages[0].shapes[0].cells;
        assert_eq!(
            cells
                .iter()
                .find(|cell| cell.name == "PinX")
                .unwrap()
                .formula
                .as_deref(),
            Some("1")
        );
        assert_eq!(
            cells
                .iter()
                .find(|cell| cell.name == "PinY")
                .unwrap()
                .formula
                .as_deref(),
            Some("1")
        );
    }

    #[test]
    fn resize_refusal_leaves_both_axes_unchanged() {
        let session = session();
        add_cell(&session, "Width", Some("1"), None);
        add_cell(&session, "Height", Some("1"), None);
        add_cell(&session, "LockHeight", Some("1"), None);
        assert!(
            session
                .resize_shape(&EditCtx::local("a"), "page:1", "page:1:shape:1", "2", "3")
                .is_err()
        );
        let cells = &session.snapshot().unwrap().pages[0].shapes[0].cells;
        assert_eq!(
            cells
                .iter()
                .find(|cell| cell.name == "Width")
                .unwrap()
                .formula
                .as_deref(),
            Some("1")
        );
        assert_eq!(
            cells
                .iter()
                .find(|cell| cell.name == "Height")
                .unwrap()
                .formula
                .as_deref(),
            Some("1")
        );
    }

    #[test]
    fn second_axis_guard_leaves_gesture_unchanged() {
        let session = session();
        add_cell(&session, "PinX", Some("1"), None);
        add_cell(&session, "PinY", Some("GUARD(1)"), None);
        assert!(
            session
                .move_shape(&EditCtx::local("a"), "page:1", "page:1:shape:1", "2", "3")
                .is_err()
        );
        let cells = &session.snapshot().unwrap().pages[0].shapes[0].cells;
        assert_eq!(
            cells
                .iter()
                .find(|cell| cell.name == "PinX")
                .unwrap()
                .formula
                .as_deref(),
            Some("1")
        );
        assert_eq!(
            cells
                .iter()
                .find(|cell| cell.name == "PinY")
                .unwrap()
                .formula
                .as_deref(),
            Some("GUARD(1)")
        );
    }

    #[test]
    fn guarded_section_row_cell_refuses_edits() {
        let session = session();
        let locator = CellLocator {
            sheet: CellSheet::Page(1),
            shape_id: Some(1),
            section: Some("Geometry".to_owned()),
            row: Some(CellRow::Index(0)),
            cell_name: "X".to_owned(),
        };
        add_cell_at(
            &session,
            "X",
            Some("Geometry"),
            Some(CellRow::Index(0)),
            Some("GUARD(1)"),
            None,
        );
        assert!(
            session
                .set_cell_formula_at(
                    &EditCtx::local("a"),
                    "page:1",
                    "page:1:shape:1",
                    locator.clone(),
                    "2"
                )
                .is_err()
        );
        let snapshot = session.snapshot().unwrap();
        let cell = snapshot.pages[0].shapes[0]
            .cells
            .iter()
            .find(|cell| cell.locator == locator)
            .unwrap();
        assert_eq!(cell.formula.as_deref(), Some("GUARD(1)"));
    }

    #[test]
    fn setatref_writes_only_the_resolved_target() {
        let session = session();
        add_cell(&session, "Width", Some("SETATREF(Target)"), None);
        add_cell(&session, "Target", Some("1"), None);
        let receipt = session
            .resize_shape(&EditCtx::local("a"), "page:1", "page:1:shape:1", "2", "3")
            .unwrap_err();
        assert!(receipt.to_string().contains("Height"));
        let receipt = session
            .set_cell_formula(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:1",
                "Width",
                "2",
            )
            .unwrap();
        assert_eq!(receipt.cell_name, "Target");
        let snapshot = session.snapshot().unwrap();
        let cells = &snapshot.pages[0].shapes[0].cells;
        assert_eq!(
            cells
                .iter()
                .find(|cell| cell.name == "Width")
                .unwrap()
                .formula
                .as_deref(),
            Some("SETATREF(Target)")
        );
        assert_eq!(
            cells
                .iter()
                .find(|cell| cell.name == "Target")
                .unwrap()
                .formula
                .as_deref(),
            Some("2")
        );
    }

    #[test]
    fn inherited_guard_materialized_in_the_crdt_refuses_edits() {
        let session = session();
        add_cell(&session, "Width", Some("GUARD(1)"), None);
        assert!(
            session
                .resize_shape(&EditCtx::local("a"), "page:1", "page:1:shape:1", "2", "3")
                .is_err()
        );
    }

    #[test]
    fn local_reorders_are_undoable() {
        let session = session();
        session
            .reorder_shape(&EditCtx::local("a"), "page:1", "page:1:shape:1", 1)
            .unwrap();
        session.add_undo_barrier();
        assert!(session.undo());
        session
            .reorder_page(&EditCtx::local("a"), "page:1", 1)
            .unwrap();
        session.add_undo_barrier();
        assert!(session.undo());
    }

    #[test]
    fn reopen_preserves_the_next_local_shape_id() {
        let session = session();
        let draft = ShapeDraft {
            source_id: 9,
            name: None,
            cells: Vec::new(),
        };
        let first = session
            .add_shape(&EditCtx::local("a"), "page:1", &draft)
            .unwrap();
        let reopened =
            DiagramSession::open_from_update(&session.encode_state_as_update_v1(), 7).unwrap();
        let second = reopened
            .add_shape(&EditCtx::local("a"), "page:1", &draft)
            .unwrap();
        assert_ne!(first.shape_id, second.shape_id);
        assert_eq!(reopened.snapshot().unwrap().pages[0].shapes.len(), 4);
    }

    #[test]
    fn add_shape_preserves_draft_section_row_cell_locators() {
        let session = session();
        let geometry = CellLocator {
            sheet: CellSheet::Page(1),
            shape_id: Some(42),
            section: Some("Geometry".to_owned()),
            row: Some(CellRow::Index(0)),
            cell_name: "X".to_owned(),
        };
        let draft = ShapeDraft {
            source_id: 42,
            name: None,
            cells: vec![
                CellSnapshot {
                    locator: geometry.clone(),
                    name: "X".to_owned(),
                    formula: Some("1".to_owned()),
                    value: None,
                },
                CellSnapshot {
                    locator: CellLocator {
                        row: Some(CellRow::Index(1)),
                        ..geometry.clone()
                    },
                    name: "X".to_owned(),
                    formula: Some("2".to_owned()),
                    value: None,
                },
            ],
        };

        let receipt = session
            .add_shape(&EditCtx::local("a"), "page:1", &draft)
            .unwrap();
        let snapshot = session.snapshot().unwrap();
        let shape = snapshot.pages[0]
            .shapes
            .iter()
            .find(|shape| shape.id == receipt.shape_id)
            .unwrap();

        assert_eq!(shape.cells.len(), 2);
        assert!(shape.cells.iter().any(|cell| {
            cell.name == "X"
                && cell.formula.as_deref() == Some("1")
                && cell.locator.section.as_deref() == Some("Geometry")
                && cell.locator.row == Some(CellRow::Index(0))
        }));
        assert!(shape.cells.iter().any(|cell| {
            cell.name == "X"
                && cell.formula.as_deref() == Some("2")
                && cell.locator.section.as_deref() == Some("Geometry")
                && cell.locator.row == Some(CellRow::Index(1))
        }));
    }

    #[test]
    fn shape_draft_round_trips_through_serde() {
        let draft = ShapeDraft {
            source_id: 42,
            name: Some("Rectangle".to_owned()),
            cells: vec![CellSnapshot {
                locator: CellLocator {
                    sheet: CellSheet::Page(1),
                    shape_id: Some(42),
                    section: Some("Geometry".to_owned()),
                    row: Some(CellRow::Name("MoveTo".to_owned())),
                    cell_name: "X".to_owned(),
                },
                name: "X".to_owned(),
                formula: Some("2".to_owned()),
                value: Some("2".to_owned()),
            }],
        };

        let serialized = serde_json::to_string(&draft).unwrap();
        assert_eq!(
            serde_json::from_str::<ShapeDraft>(&serialized).unwrap(),
            draft
        );
    }

    #[test]
    fn remote_protected_formula_rewrite_is_rejected_but_legitimate_update_is_accepted() {
        let session = session();
        add_cell(&session, "Width", Some("GUARD(1)"), None);
        let attacker = doc_with_client_id(9);
        hydrate_doc(&attacker, &session.encode_state_as_update_v1()).unwrap();
        let mut txn = attacker.transact_mut_with(9_u64);
        let sheets = txn.get_map(SHEETS).unwrap();
        let shape = match sheets.get(&txn, "page:1:shape:1") {
            Some(yrs::Out::YMap(shape)) => shape,
            _ => unreachable!(),
        };
        let cells = match shape.get(&txn, "cells") {
            Some(yrs::Out::YMap(cells)) => cells,
            _ => unreachable!(),
        };
        let width = match cells.get(&txn, "Width") {
            Some(yrs::Out::YMap(cell)) => cell,
            _ => unreachable!(),
        };
        width.insert(&mut txn, "formula", "2");
        drop(txn);
        let update = attacker
            .transact()
            .encode_diff_v1(&session.doc.transact().state_vector());
        assert!(session.apply_update_v1(&update).is_err());
        assert_eq!(
            session.snapshot().unwrap().pages[0].shapes[0]
                .cells
                .iter()
                .find(|cell| cell.name == "Width")
                .unwrap()
                .formula
                .as_deref(),
            Some("GUARD(1)")
        );
        let legitimate = doc_with_client_id(10);
        hydrate_doc(&legitimate, &session.encode_state_as_update_v1()).unwrap();
        let mut txn = legitimate.transact_mut_with(10_u64);
        let sheets = txn.get_map(SHEETS).unwrap();
        let shape = match sheets.get(&txn, "page:1:shape:1") {
            Some(yrs::Out::YMap(shape)) => shape,
            _ => unreachable!(),
        };
        let cells = match shape.get(&txn, "cells") {
            Some(yrs::Out::YMap(cells)) => cells,
            _ => unreachable!(),
        };
        let cell = cells.insert(&mut txn, "PinX", MapPrelim::default());
        cell.insert(&mut txn, "name", "PinX");
        cell.insert(&mut txn, "formula", "2");
        drop(txn);
        let update = legitimate
            .transact()
            .encode_diff_v1(&session.doc.transact().state_vector());
        assert!(session.apply_update_v1(&update).is_ok());
    }

    #[test]
    fn remote_rewrite_of_a_nested_child_guard_is_rejected() {
        let session = session();
        let child = "page:1:shape:1:shape:3";
        add_child_shape(&session, child, "page:1:shape:1");
        add_shape_cell(&session, child, "Width", None, None, Some("GUARD(1)"), None);
        let peer = peer_doc(&session, 9);
        write_peer_cell_field(&peer, child, "Width", "formula", "2");
        assert!(
            session
                .apply_update_v1(&peer_update(&session, &peer))
                .is_err()
        );
        let snapshot = session.snapshot().unwrap();
        assert_eq!(
            snapshot.pages[0].shapes[0].children[0]
                .cells
                .iter()
                .find(|cell| cell.name == "Width")
                .unwrap()
                .formula
                .as_deref(),
            Some("GUARD(1)")
        );
    }

    #[test]
    fn remote_rewrite_of_a_nested_child_locked_cell_is_rejected() {
        let session = session();
        let child = "page:1:shape:1:shape:3";
        add_child_shape(&session, child, "page:1:shape:1");
        add_shape_cell(&session, child, "Width", None, None, Some("1"), None);
        add_shape_cell(&session, child, "LockWidth", None, None, Some("1"), None);
        let peer = peer_doc(&session, 9);
        write_peer_cell_field(&peer, child, "Width", "formula", "5");
        assert!(
            session
                .apply_update_v1(&peer_update(&session, &peer))
                .is_err()
        );
        assert_eq!(
            session.snapshot().unwrap().pages[0].shapes[0].children[0]
                .cells
                .iter()
                .find(|cell| cell.name == "Width")
                .unwrap()
                .formula
                .as_deref(),
            Some("1")
        );
    }

    #[test]
    fn remote_baseline_forgery_cannot_suppress_a_collaborative_edit() {
        let session = session();
        add_cell(&session, "Width", Some("1"), None);
        session
            .set_cell_formula(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:1",
                "Width",
                "2",
            )
            .unwrap();
        let peer = peer_doc(&session, 9);
        write_peer_cell_field(&peer, "page:1:shape:1", "Width", "baselineFormula", "2");
        assert!(
            session
                .apply_update_v1(&peer_update(&session, &peer))
                .is_err()
        );
        assert!(session.semantic_cell_edits().unwrap().iter().any(|edit| {
            edit.locator.cell_name == "Width" && edit.formula.as_deref() == Some("2")
        }));
    }

    #[test]
    fn remote_rewrites_of_shape_identity_are_rejected() {
        for (field, value) in [
            ("origin", "session"),
            ("pageId", "page:2"),
            ("parentId", "page:1:shape:2"),
        ] {
            let session = session();
            let peer = peer_doc(&session, 9);
            write_peer_shape_field(&peer, "page:1:shape:1", field, value);
            assert!(
                session
                    .apply_update_v1(&peer_update(&session, &peer))
                    .is_err(),
                "{field}"
            );
        }
    }

    #[test]
    fn remote_new_shape_cannot_forge_package_provenance() {
        let session = session();
        let before = session.encode_state_as_update_v1();
        let peer = peer_doc(&session, 9);
        write_peer_new_shape(&peer, "page:1:shape:forged", "page:1", "package", 1.0);
        write_peer_new_cell(&peer, "page:1:shape:forged", "Width", "999");
        assert!(
            session
                .apply_update_v1(&peer_update(&session, &peer))
                .is_err()
        );
        assert_eq!(before, session.encode_state_as_update_v1());
    }

    #[test]
    fn state_vectors_are_limited_before_decode() {
        assert!(decode_state_vector_v1(&vec![0; MAX_STATE_VECTOR_BYTES + 1]).is_err());
    }

    #[test]
    fn peers_converge_after_exchanging_updates() {
        let seed = session();
        add_cell(&seed, "Width", Some("1"), None);
        add_cell(&seed, "PinX", Some("1"), None);
        let state = seed.encode_state_as_update_v1();
        let left = DiagramSession::open_from_update(&state, 11).unwrap();
        let right = DiagramSession::open_from_update(&state, 12).unwrap();
        left.set_cell_formula(
            &EditCtx::local("left"),
            "page:1",
            "page:1:shape:1",
            "Width",
            "2",
        )
        .unwrap();
        right
            .set_cell_formula(
                &EditCtx::local("right"),
                "page:1",
                "page:1:shape:1",
                "PinX",
                "3",
            )
            .unwrap();
        let left_update = left
            .encode_diff_v1(&right.encode_state_vector_v1())
            .unwrap();
        let right_update = right
            .encode_diff_v1(&left.encode_state_vector_v1())
            .unwrap();
        left.apply_update_v1(&right_update).unwrap();
        right.apply_update_v1(&left_update).unwrap();
        assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    }

    #[test]
    fn peers_converge_after_editing_distinct_section_row_cells() {
        let seed = session();
        let x = CellLocator {
            sheet: CellSheet::Page(1),
            shape_id: Some(1),
            section: Some("Geometry".to_owned()),
            row: Some(CellRow::Index(0)),
            cell_name: "X".to_owned(),
        };
        let y = CellLocator {
            cell_name: "Y".to_owned(),
            ..x.clone()
        };
        add_cell_at(
            &seed,
            "X",
            Some("Geometry"),
            Some(CellRow::Index(0)),
            Some("1"),
            None,
        );
        add_cell_at(
            &seed,
            "Y",
            Some("Geometry"),
            Some(CellRow::Index(0)),
            Some("1"),
            None,
        );
        let state = seed.encode_state_as_update_v1();
        let left = DiagramSession::open_from_update(&state, 11).unwrap();
        let right = DiagramSession::open_from_update(&state, 12).unwrap();
        left.set_cell_formula_at(&EditCtx::local("left"), "page:1", "page:1:shape:1", x, "2")
            .unwrap();
        right
            .set_cell_formula_at(&EditCtx::local("right"), "page:1", "page:1:shape:1", y, "3")
            .unwrap();
        let left_update = left
            .encode_diff_v1(&right.encode_state_vector_v1())
            .unwrap();
        let right_update = right
            .encode_diff_v1(&left.encode_state_vector_v1())
            .unwrap();
        left.apply_update_v1(&right_update).unwrap();
        right.apply_update_v1(&left_update).unwrap();
        let snapshot = left.snapshot().unwrap();
        assert_eq!(snapshot, right.snapshot().unwrap());
        assert!(
            snapshot.pages[0].shapes[0]
                .cells
                .iter()
                .any(|cell| { cell.name == "X" && cell.formula.as_deref() == Some("2") })
        );
        assert!(
            snapshot.pages[0].shapes[0]
                .cells
                .iter()
                .any(|cell| { cell.name == "Y" && cell.formula.as_deref() == Some("3") })
        );
    }

    #[test]
    fn grouped_child_cells_are_addressable_from_snapshots() {
        let session = DiagramSession::open(
            include_bytes!("../../vsdx-parse/tests/fixtures/nested-groups.vsdx"),
            7,
        )
        .unwrap();
        let parent = &session.snapshot().unwrap().pages[0].shapes[0];
        let child = parent.children.first().unwrap();
        let cell = child.cells.first().unwrap();
        assert_eq!(cell.locator.shape_id, Some(child.source_id));
        session
            .set_cell_formula_at(
                &EditCtx::local("a"),
                "page:1",
                &child.id,
                cell.locator.clone(),
                "42",
            )
            .unwrap();
        let snapshot = session.snapshot().unwrap();
        let changed = snapshot.pages[0].shapes[0].children[0]
            .cells
            .iter()
            .find(|candidate| candidate.locator == cell.locator)
            .unwrap();
        assert_eq!(changed.formula.as_deref(), Some("42"));
    }

    #[test]
    fn shared_seed_is_identical_for_distinct_clients() {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx");
        let first = DiagramSession::open(source, 17).unwrap();
        let second = DiagramSession::open(source, 18).unwrap();
        assert_eq!(
            first.encode_state_as_update_v1(),
            second.encode_state_as_update_v1()
        );
    }

    #[test]
    fn undo_keeps_remote_edits() {
        let seed = session();
        add_cell(&seed, "Width", Some("1"), None);
        add_cell(&seed, "PinX", Some("1"), None);
        let state = seed.encode_state_as_update_v1();
        let local = DiagramSession::open_from_update(&state, 21).unwrap();
        let remote = DiagramSession::open_from_update(&state, 22).unwrap();
        local
            .set_cell_formula(
                &EditCtx::local("local"),
                "page:1",
                "page:1:shape:1",
                "Width",
                "2",
            )
            .unwrap();
        local.add_undo_barrier();
        remote
            .set_cell_formula(
                &EditCtx::local("remote"),
                "page:1",
                "page:1:shape:1",
                "PinX",
                "3",
            )
            .unwrap();
        local
            .apply_update_v1(
                &remote
                    .encode_diff_v1(&local.encode_state_vector_v1())
                    .unwrap(),
            )
            .unwrap();
        assert!(local.undo());
        let cells = &local.snapshot().unwrap().pages[0].shapes[0].cells;
        assert_eq!(
            cells
                .iter()
                .find(|cell| cell.name == "Width")
                .unwrap()
                .formula
                .as_deref(),
            Some("1")
        );
        assert_eq!(
            cells
                .iter()
                .find(|cell| cell.name == "PinX")
                .unwrap()
                .formula
                .as_deref(),
            Some("3")
        );
    }

    #[test]
    fn malformed_updates_and_vectors_leave_the_document_unchanged() {
        let session = session();
        let before = session.encode_state_as_update_v1();
        let mut trailing = before.clone();
        trailing.push(0);
        assert!(session.apply_update_v1(&trailing).is_err());
        assert!(
            session
                .apply_update_v1(&vec![0; MAX_UPDATE_BYTES + 1])
                .is_err()
        );
        assert!(session.encode_diff_v1(&[0, 0]).is_err());
        assert!(
            session
                .encode_diff_v1(&vec![0; MAX_STATE_VECTOR_BYTES + 1])
                .is_err()
        );
        assert_eq!(before, session.encode_state_as_update_v1());
    }
}
