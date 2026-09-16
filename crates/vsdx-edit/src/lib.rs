//! Collaborative yrs-backed VSDX diagram model.

use std::cell::RefCell;
use std::panic::{AssertUnwindSafe, catch_unwind};

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

#[cfg(feature = "wasm")]
pub mod wasm;

pub use model::*;
pub use undo::DiagramUndoManager;

pub(crate) const META: &str = "vsdx:meta";
pub(crate) const PAGE_ORDER: &str = "vsdx:page-order";
pub(crate) const PAGES: &str = "vsdx:pages";
pub(crate) const SHEETS: &str = "vsdx:sheets";
pub(crate) const CONNECTS: &str = "vsdx:connects";
pub(crate) const STORIES: &str = "vsdx:stories";
pub(crate) const REMOTE_ORIGIN: &str = "vsdx:remote";
pub(crate) const HYDRATE_ORIGIN: &str = "vsdx:hydrate";
pub(crate) const MIGRATE_ORIGIN: &str = "vsdx:migrate";
const BOOTSTRAP_CLIENT_ID: u64 = (1_u64 << 53) - 1;
pub const MAX_SAFE_CLIENT_ID: u64 = BOOTSTRAP_CLIENT_ID - 1;
pub const MAX_UPDATE_BYTES: usize = 64 * 1024 * 1024;
const MAX_STATE_VECTOR_ENTRIES: u32 = 65_536;
const MAX_STATE_VECTOR_BYTES: usize = 1024 * 1024;

/// A collaborative session with private CRDT storage.
pub struct DiagramSession {
    pub(crate) doc: Doc,
    client_id: u64,
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
        Ok(Self {
            doc,
            client_id,
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
        diagram::migrate_doc(&doc)?;
        diagram::validate_doc(&doc)?;
        let undo = DiagramUndoManager::new(&doc, client_id)?;
        Ok(Self {
            doc,
            client_id,
            undo: RefCell::new(undo),
        })
    }

    pub fn client_id(&self) -> u64 {
        self.client_id
    }
    pub fn package(&self) -> EditResult<vsdx_parse::VsdxPackage> {
        diagram::package_from_doc(&self.doc)
    }
    #[cfg(test)]
    pub(crate) fn yrs_doc(&self) -> &Doc {
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
        diagram::normalize_concurrent_orders(&staged)?;
        diagram::validate_remote_update(&self.doc, &staged)?;
        let update = staged
            .transact()
            .encode_state_as_update_v1(&self.doc.transact().state_vector());
        self.doc
            .transact_mut_with(REMOTE_ORIGIN)
            .apply_update(decode_update_v1(&update).map_err(EditError::InvalidUpdate)?)
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
    for root in [META, PAGES, SHEETS, CONNECTS, STORIES] {
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
        meta.insert(&mut txn, "schemaVersion", 2.0);
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
        txn.get_or_insert_map(CONNECTS);
        for page_id in ["page:1", "page:2"] {
            order.push_back(&mut txn, page_id);
            let page = pages.insert(&mut txn, page_id, MapPrelim::default());
            page.insert(&mut txn, "id", page_id);
            page.insert(&mut txn, "sourcePartPath", format!("/{page_id}"));
            page.insert(&mut txn, "maxSourceId", 2.0);
            page.insert(&mut txn, "shapes", ArrayPrelim::default());
        }
        for (id, page_id, source_id) in [
            ("page:1:shape:1", "page:1", 1.0),
            ("page:1:shape:2", "page:1", 2.0),
        ] {
            let shape = sheets.insert(&mut txn, id, MapPrelim::default());
            shape.insert(&mut txn, "id", id);
            shape.insert(&mut txn, "pageId", page_id);
            shape.insert(&mut txn, "sourceId", source_id);
            shape.insert(&mut txn, "origin", "original");
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
        let mut txn = session.doc.transact_mut_with(HYDRATE_ORIGIN);
        let sheets = txn.get_map(SHEETS).unwrap();
        let shape = match sheets.get(&txn, "page:1:shape:1") {
            Some(yrs::Out::YMap(shape)) => shape,
            _ => unreachable!(),
        };
        let cells = match shape.get(&txn, "cells") {
            Some(yrs::Out::YMap(cells)) => cells,
            _ => unreachable!(),
        };
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
        }
        if let Some(value) = value {
            cell.insert(&mut txn, "value", value);
        }
    }

    fn add_child_shape(session: &DiagramSession, id: &str, parent_id: &str) {
        let mut txn = session.doc.transact_mut_with(HYDRATE_ORIGIN);
        let sheets = txn.get_map(SHEETS).unwrap();
        let shape = sheets.insert(&mut txn, id, MapPrelim::default());
        shape.insert(&mut txn, "id", id);
        shape.insert(&mut txn, "pageId", "page:1");
        shape.insert(&mut txn, "origin", "original");
        shape.insert(&mut txn, "sourceId", 3.0);
        shape.insert(&mut txn, "parentId", parent_id);
        shape.insert(&mut txn, "cells", MapPrelim::default());
        shape.insert(&mut txn, "shapes", ArrayPrelim::default());
        // The parent may not exist yet (a forward reference, as when a test wires up a mutual
        // cycle); attaching to its child order then happens once it does, via `attach_child`.
        let Some(yrs::Out::YMap(parent)) = sheets.get(&txn, parent_id) else {
            return;
        };
        let child_order = match parent.get(&txn, "shapes") {
            Some(yrs::Out::YArray(child_order)) => child_order,
            _ => parent.insert(&mut txn, "shapes", ArrayPrelim::default()),
        };
        child_order.push_back(&mut txn, id);
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
    fn snapshots_evaluate_trusted_formulas_with_qualified_section_references() {
        let session = session();
        add_shape_cell(
            &session,
            "page:1:shape:1",
            "Width",
            None,
            None,
            Some("=4"),
            None,
        );
        add_shape_cell(
            &session,
            "page:1:shape:1",
            "Value",
            Some("User"),
            Some(CellRow::Name("Scale".into())),
            Some("3"),
            None,
        );
        add_shape_cell(
            &session,
            "page:1:shape:1",
            "PinX",
            None,
            None,
            Some("Width/2+User.Scale"),
            None,
        );
        let snapshot = session.snapshot().unwrap();
        let pin = snapshot.pages[0].shapes[0]
            .cells
            .iter()
            .find(|cell| cell.name == "PinX")
            .unwrap();
        assert_eq!(pin.value.as_deref(), Some("5"));
        session
            .set_cell_formula(
                &EditCtx::local("test"),
                "page:1",
                "page:1:shape:1",
                "PinX",
                "Width*User.Scale",
            )
            .unwrap();
        let edits = session.semantic_cell_edits().unwrap();
        assert_eq!(edits.len(), 1);
        assert_eq!(edits[0].value.as_deref(), Some("12"));
    }

    #[test]
    fn snapshot_recomputes_dependent_values_after_editing_their_inputs() {
        let session = session();
        add_shape_cell(
            &session,
            "page:1:shape:1",
            "Width",
            None,
            None,
            Some("2"),
            Some("2"),
        );
        add_shape_cell(
            &session,
            "page:1:shape:1",
            "PinX",
            None,
            None,
            Some("Width/2"),
            Some("1"),
        );
        session
            .set_cell_formula(
                &EditCtx::local("test"),
                "page:1",
                "page:1:shape:1",
                "Width",
                "4",
            )
            .unwrap();
        let snapshot = session.snapshot().unwrap();
        let pin = snapshot.pages[0].shapes[0]
            .cells
            .iter()
            .find(|cell| cell.name == "PinX")
            .unwrap();
        assert_eq!(pin.value.as_deref(), Some("2"));
    }

    #[test]
    fn drafts_fail_atomically_for_invalid_or_duplicate_locators() {
        let session = session();
        let cell = CellSnapshot {
            locator: CellLocator {
                sheet: CellSheet::Page(1),
                shape_id: None,
                section: Some("Geometry".into()),
                section_index: None,
                row: Some(CellRow::Index(0)),
                cell_name: "X".into(),
            },
            row_type: Some("MoveTo".into()),
            name: "X".into(),
            formula: Some("1".into()),
            value: None,
        };
        let mut duplicate = cell.clone();
        duplicate.locator.section_index = Some(0);
        let mut row_without_section = cell.clone();
        row_without_section.locator.section = None;
        let mut invalid_formula = cell.clone();
        invalid_formula.formula = Some("1+".into());
        let mut invalid_xml = cell.clone();
        invalid_xml.name.push('\0');
        invalid_xml.locator.cell_name = invalid_xml.name.clone();
        let mut raw_cache = cell.clone();
        raw_cache.value = Some("999".into());
        for cells in [
            vec![cell.clone(), duplicate],
            vec![row_without_section],
            vec![invalid_formula],
            vec![invalid_xml],
            vec![raw_cache],
        ] {
            let before = session.encode_state_as_update_v1();
            assert!(
                session
                    .add_shape(
                        &EditCtx::local("test"),
                        "page:1",
                        &ShapeDraft { name: None, cells }
                    )
                    .is_err()
            );
            assert_eq!(session.encode_state_as_update_v1(), before);
        }
    }

    #[test]
    fn added_shape_ids_are_not_reused_after_deletion() {
        let session = session();
        let draft = ShapeDraft {
            name: None,
            cells: Vec::new(),
        };
        let first = session
            .add_shape(&EditCtx::local("test"), "page:1", &draft)
            .unwrap();
        session
            .delete_shape(&EditCtx::local("test"), "page:1", &first.shape_id)
            .unwrap();
        let reopened =
            DiagramSession::open_from_update(&session.encode_state_as_update_v1(), 7).unwrap();
        let second = reopened
            .add_shape(&EditCtx::local("test"), "page:1", &draft)
            .unwrap();
        assert_ne!(first.shape_id, second.shape_id);
    }

    #[test]
    fn hydrated_numeric_identities_reject_invalid_numbers() {
        for field in ["sectionIndex", "rowIndex", "sourceId", "maxSourceId"] {
            for value in [
                -1.0,
                0.5,
                f64::NAN,
                f64::INFINITY,
                f64::from(u32::MAX) + 1.0,
            ] {
                let session = session();
                add_cell_at(
                    &session,
                    "X",
                    Some("Geometry"),
                    Some(CellRow::Index(0)),
                    Some("1"),
                    None,
                );
                let peer = peer_doc(&session, 9);
                let mut txn = peer.transact_mut();
                let owner = if field == "maxSourceId" {
                    match txn.get_map(PAGES).unwrap().get(&txn, "page:1").unwrap() {
                        yrs::Out::YMap(page) => page,
                        _ => unreachable!(),
                    }
                } else if field == "sourceId" {
                    match txn
                        .get_map(SHEETS)
                        .unwrap()
                        .get(&txn, "page:1:shape:1")
                        .unwrap()
                    {
                        yrs::Out::YMap(shape) => shape,
                        _ => unreachable!(),
                    }
                } else {
                    match shape_cells(&txn, "page:1:shape:1")
                        .get(&txn, "Geometry\u{1f}IX:0\u{1f}X")
                        .unwrap()
                    {
                        yrs::Out::YMap(cell) => cell,
                        _ => unreachable!(),
                    }
                };
                owner.insert(&mut txn, field, value);
                drop(txn);
                let update = peer
                    .transact()
                    .encode_state_as_update_v1(&StateVector::default());
                assert!(
                    DiagramSession::open_from_update(&update, 10).is_err(),
                    "{field}={value}"
                );
            }
        }
    }

    #[test]
    fn concurrent_reorders_of_the_same_shape_converge() {
        let left = session();
        let right = DiagramSession::open_from_update(&left.encode_state_as_update_v1(), 8).unwrap();
        left.reorder_shape(&EditCtx::local("left"), "page:1", "page:1:shape:1", 1)
            .unwrap();
        right
            .reorder_shape(&EditCtx::local("right"), "page:1", "page:1:shape:1", 1)
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
        assert_eq!(left.snapshot().unwrap().pages[0].shapes.len(), 2);
    }

    #[test]
    fn concurrent_page_reorders_and_shape_deletion_converge() {
        for delete in [false, true] {
            let left = session();
            let right =
                DiagramSession::open_from_update(&left.encode_state_as_update_v1(), 8).unwrap();
            if delete {
                left.reorder_shape(&EditCtx::local("left"), "page:1", "page:1:shape:1", 1)
                    .unwrap();
                right
                    .delete_shape(&EditCtx::local("right"), "page:1", "page:1:shape:1")
                    .unwrap();
            } else {
                left.reorder_page(&EditCtx::local("left"), "page:1", 1)
                    .unwrap();
                right
                    .reorder_page(&EditCtx::local("right"), "page:1", 1)
                    .unwrap();
            }
            let left_update = left
                .encode_diff_v1(&right.encode_state_vector_v1())
                .unwrap();
            let right_update = right
                .encode_diff_v1(&left.encode_state_vector_v1())
                .unwrap();
            left.apply_update_v1(&right_update).unwrap();
            right.apply_update_v1(&left_update).unwrap();
            assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
            let left_update = left
                .encode_diff_v1(&right.encode_state_vector_v1())
                .unwrap();
            let right_update = right
                .encode_diff_v1(&left.encode_state_vector_v1())
                .unwrap();
            left.apply_update_v1(&right_update).unwrap();
            right.apply_update_v1(&left_update).unwrap();
            assert_eq!(
                left.encode_state_vector_v1(),
                right.encode_state_vector_v1()
            );
            assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
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
    fn shape_bounds_refusal_preserves_all_cells_and_emits_no_update() {
        for (name, formula) in [
            ("PinX", "GUARD(1)"),
            ("LockMoveY", "1"),
            ("LockHeight", "1"),
        ] {
            let session = session();
            for cell in ["PinX", "PinY", "Width", "Height"] {
                add_cell(&session, cell, Some("1"), None);
            }
            add_cell(&session, name, Some(formula), None);
            let before = session.snapshot().unwrap();
            let vector = session.encode_state_vector_v1();
            assert!(
                session
                    .set_shape_bounds(
                        &EditCtx::local("a"),
                        "page:1",
                        "page:1:shape:1",
                        ["2", "3", "4", "5"].map(str::to_owned)
                    )
                    .is_err()
            );
            assert_eq!(session.snapshot().unwrap(), before);
            assert_eq!(session.encode_state_vector_v1(), vector);
        }
    }

    #[test]
    fn shape_bounds_undo_restores_all_four_cells() {
        let session = session();
        for cell in ["PinX", "PinY", "Width", "Height"] {
            add_cell(&session, cell, Some("1"), None);
        }
        let before = session.snapshot().unwrap();
        let receipts = session
            .set_shape_bounds(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:1",
                ["2", "3", "4", "5"].map(str::to_owned),
            )
            .unwrap();
        assert_eq!(receipts.map(|receipt| receipt.after), ["2", "3", "4", "5"]);
        assert!(session.undo());
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(!session.can_undo());
    }

    #[test]
    fn resize_loc_pin_evaluates_formulas_without_mutating() {
        let session = session();
        for (name, formula) in [
            ("Width", "2"),
            ("Height", "3"),
            ("LocPinX", "Width*0.5+0.25"),
            ("LocPinY", "0.75"),
        ] {
            add_cell(&session, name, Some(formula), None);
        }
        let before = session.snapshot().unwrap();
        assert_eq!(
            session
                .resize_loc_pin("page:1", "page:1:shape:1", 4.0, 6.0)
                .unwrap(),
            [2.25, 0.75]
        );
        assert_eq!(session.snapshot().unwrap(), before);
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

    fn container_pair() -> DiagramSession {
        let session = session();
        for (shape, pin_x, pin_y, width, height, loc_pin) in [
            ("page:1:shape:1", "5", "4", "4", "4", "2"),
            ("page:1:shape:2", "5", "4", "1", "1", "0.5"),
        ] {
            for (name, formula) in [
                ("PinX", pin_x),
                ("PinY", pin_y),
                ("Width", width),
                ("Height", height),
                ("LocPinX", loc_pin),
                ("LocPinY", loc_pin),
            ] {
                add_shape_cell(&session, shape, name, None, None, Some(formula), None);
            }
        }
        add_shape_cell(
            &session,
            "page:1:shape:1",
            "Relationships",
            None,
            None,
            Some("SUM(DEPENDSON(1,Sheet.2!SheetRef()))"),
            None,
        );
        add_shape_cell(
            &session,
            "page:1:shape:1",
            "Value",
            Some("User"),
            Some(CellRow::Name("msvStructureType".into())),
            None,
            Some("Container"),
        );
        add_shape_cell(
            &session,
            "page:1:shape:1",
            "Value",
            Some("User"),
            Some(CellRow::Name("msvSDContainerMargin".into())),
            None,
            Some("0.25"),
        );
        add_shape_cell(
            &session,
            "page:1:shape:2",
            "Relationships",
            None,
            None,
            Some("SUM(DEPENDSON(4,Sheet.1!SheetRef()))"),
            None,
        );
        session
    }

    fn shape_formula(session: &DiagramSession, shape: &str, name: &str) -> Option<String> {
        session.snapshot().unwrap().pages[0]
            .shapes
            .iter()
            .find(|candidate| candidate.id == shape)
            .and_then(|shape| {
                shape
                    .cells
                    .iter()
                    .find(|cell| cell.name == name)
                    .and_then(|cell| cell.formula.clone())
            })
    }

    #[test]
    fn container_move_shifts_members_in_one_undo() {
        let session = container_pair();
        let before = session.snapshot().unwrap();
        let receipts = session
            .move_container(&EditCtx::local("a"), "page:1", "page:1:shape:1", 1.0, 2.0)
            .unwrap();
        assert_eq!(receipts.len(), 4);
        assert_eq!(
            shape_formula(&session, "page:1:shape:1", "PinX").as_deref(),
            Some("6")
        );
        assert_eq!(
            shape_formula(&session, "page:1:shape:1", "PinY").as_deref(),
            Some("6")
        );
        assert_eq!(
            shape_formula(&session, "page:1:shape:2", "PinX").as_deref(),
            Some("6")
        );
        assert_eq!(
            shape_formula(&session, "page:1:shape:2", "PinY").as_deref(),
            Some("6")
        );
        assert!(session.undo());
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(!session.can_undo());
    }

    #[test]
    fn container_move_refusal_moves_nothing() {
        let session = container_pair();
        add_shape_cell(
            &session,
            "page:1:shape:2",
            "LockMoveX",
            None,
            None,
            Some("1"),
            None,
        );
        let before = session.snapshot().unwrap();
        assert!(
            session
                .move_container(&EditCtx::local("a"), "page:1", "page:1:shape:1", 1.0, 2.0)
                .is_err()
        );
        assert_eq!(session.snapshot().unwrap(), before);
    }

    #[test]
    fn container_move_rejects_non_containers() {
        let session = container_pair();
        assert!(
            session
                .move_container(&EditCtx::local("a"), "page:1", "page:1:shape:2", 1.0, 2.0)
                .is_err()
        );
    }

    #[test]
    fn member_move_expands_container_with_margin() {
        let session = container_pair();
        let receipts = session
            .move_container_member(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:2",
                "8".to_owned(),
                "4".to_owned(),
            )
            .unwrap();
        assert_eq!(
            shape_formula(&session, "page:1:shape:2", "PinX").as_deref(),
            Some("8")
        );
        assert_eq!(
            shape_formula(&session, "page:1:shape:1", "Width").as_deref(),
            Some("5.75")
        );
        assert_eq!(
            shape_formula(&session, "page:1:shape:1", "PinX").as_deref(),
            Some("5")
        );
        assert!(receipts.len() >= 3);
        assert!(session.undo());
        assert_eq!(
            shape_formula(&session, "page:1:shape:1", "Width").as_deref(),
            Some("4")
        );
        assert!(!session.can_undo());
    }

    #[test]
    fn member_move_inside_leaves_container_bounds() {
        let session = container_pair();
        let receipts = session
            .move_container_member(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:2",
                "5.5".to_owned(),
                "4".to_owned(),
            )
            .unwrap();
        assert_eq!(receipts.len(), 2);
        assert_eq!(
            shape_formula(&session, "page:1:shape:1", "Width").as_deref(),
            Some("4")
        );
    }

    #[test]
    fn locked_container_skips_autofit() {
        let session = container_pair();
        add_shape_cell(
            &session,
            "page:1:shape:1",
            "Value",
            Some("User"),
            Some(CellRow::Name("msvSDContainerLocked".into())),
            None,
            Some("1"),
        );
        let receipts = session
            .move_container_member(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:2",
                "8".to_owned(),
                "4".to_owned(),
            )
            .unwrap();
        assert_eq!(receipts.len(), 2);
        assert_eq!(
            shape_formula(&session, "page:1:shape:1", "Width").as_deref(),
            Some("4")
        );
    }

    #[test]
    fn autofit_container_shrinks_to_member_extent_with_margin() {
        let session = container_pair();
        session
            .autofit_container(&EditCtx::local("a"), "page:1", "page:1:shape:1")
            .unwrap();
        assert_eq!(
            shape_formula(&session, "page:1:shape:1", "Width").as_deref(),
            Some("1.5")
        );
        assert_eq!(
            shape_formula(&session, "page:1:shape:1", "Height").as_deref(),
            Some("1.5")
        );
    }

    #[test]
    fn autofit_container_uses_formula_loc_pin_at_the_requested_size() {
        let session = container_pair();
        add_shape_cell(
            &session,
            "page:1:shape:1",
            "LocPinX",
            None,
            None,
            Some("Width*0.5"),
            None,
        );
        add_shape_cell(
            &session,
            "page:1:shape:1",
            "LocPinY",
            None,
            None,
            Some("Height*0.5"),
            None,
        );
        session
            .autofit_container(&EditCtx::local("a"), "page:1", "page:1:shape:1")
            .unwrap();
        assert_eq!(
            shape_formula(&session, "page:1:shape:1", "PinX").as_deref(),
            Some("5")
        );
        assert_eq!(
            shape_formula(&session, "page:1:shape:1", "PinY").as_deref(),
            Some("4")
        );
    }

    #[test]
    fn rotated_members_are_enclosed_by_autofit_and_move() {
        let session = container_pair();
        add_shape_cell(
            &session,
            "page:1:shape:2",
            "Angle",
            None,
            None,
            Some("0.7853981633974483"),
            None,
        );
        session
            .autofit_container(&EditCtx::local("a"), "page:1", "page:1:shape:1")
            .unwrap();
        assert_eq!(
            shape_formula(&session, "page:1:shape:1", "Width").as_deref(),
            Some("1.9142135623730958"),
        );

        let session = container_pair();
        add_shape_cell(
            &session,
            "page:1:shape:2",
            "Angle",
            None,
            None,
            Some("0.7853981633974483"),
            None,
        );
        session
            .move_container_member(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:2",
                "8".to_owned(),
                "4".to_owned(),
            )
            .unwrap();
        let width = shape_formula(&session, "page:1:shape:1", "Width")
            .unwrap()
            .parse::<f64>()
            .unwrap();
        assert!(width > 5.9);
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
            section_index: None,
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
            section_index: None,
            row: Some(CellRow::Index(0)),
            cell_name: "X".to_owned(),
        };
        let draft = ShapeDraft {
            name: None,
            cells: vec![
                CellSnapshot {
                    row_type: None,
                    locator: geometry.clone(),
                    name: "X".to_owned(),
                    formula: Some("1".to_owned()),
                    value: None,
                },
                CellSnapshot {
                    row_type: None,
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

    /// Excludes draft cells from original-package edits.
    #[test]
    fn session_added_shapes_do_not_leak_into_semantic_cell_edits() {
        let session = session();
        add_cell(&session, "Height", Some("1"), None);
        let receipt = session
            .add_shape(
                &EditCtx::local("a"),
                "page:1",
                &ShapeDraft {
                    name: Some("Added".to_owned()),
                    cells: vec![CellSnapshot {
                        row_type: None,
                        locator: CellLocator {
                            sheet: CellSheet::Page(1),
                            shape_id: None,
                            section: None,
                            section_index: None,
                            row: None,
                            cell_name: "Width".to_owned(),
                        },
                        name: "Width".to_owned(),
                        formula: Some("5".to_owned()),
                        value: None,
                    }],
                },
            )
            .unwrap();
        // Edited after creation so its formula diverges from its own draft baseline: with the
        // `ShapeOrigin::Original` filter removed, this is exactly the shape this loop would
        // otherwise still have a reason to visit.
        session
            .set_cell_formula(
                &EditCtx::local("a"),
                "page:1",
                &receipt.shape_id,
                "Width",
                "9",
            )
            .unwrap();
        session
            .set_cell_formula(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:1",
                "Height",
                "2",
            )
            .unwrap();
        let edits = session.semantic_cell_edits().unwrap();
        assert!(
            edits.iter().all(|edit| edit.locator.shape_id != Some(0)),
            "an added shape's own draft cells must never surface as package-wide semantic cell edits: {edits:?}"
        );
        assert!(
            edits
                .iter()
                .any(|edit| edit.locator.shape_id == Some(1) && edit.locator.cell_name == "Height"),
            "the original shape's own edit must still surface: {edits:?}"
        );
    }

    #[test]
    fn shape_draft_round_trips_through_serde() {
        let draft = ShapeDraft {
            name: Some("Rectangle".to_owned()),
            cells: vec![CellSnapshot {
                row_type: None,
                locator: CellLocator {
                    sheet: CellSheet::Page(1),
                    shape_id: Some(42),
                    section: Some("Geometry".to_owned()),
                    section_index: None,
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
    fn remote_setatref_formula_rewrite_is_rejected_without_changing_the_document() {
        let session = session();
        add_cell(&session, "Width", Some("SETATREF(Target)"), None);
        add_cell(&session, "Target", Some("1"), None);
        let before = session.encode_state_as_update_v1();
        let attacker = doc_with_client_id(9);
        hydrate_doc(&attacker, &before).unwrap();
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
        assert_eq!(before, session.encode_state_as_update_v1());
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
            ("origin", "added"),
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
        write_peer_new_shape(&peer, "page:1:shape:forged", "page:1", "original", 1.0);
        write_peer_new_cell(&peer, "page:1:shape:forged", "Width", "999");
        assert!(
            session
                .apply_update_v1(&peer_update(&session, &peer))
                .is_err()
        );
        assert_eq!(before, session.encode_state_as_update_v1());
    }

    #[test]
    fn remote_new_shape_with_a_self_parent_cycle_is_rejected() {
        let session = session();
        let before = session.encode_state_as_update_v1();
        let peer = peer_doc(&session, 9);
        write_peer_new_shape(&peer, "page:1:shape:forged", "page:1", "session", 1.0);
        write_peer_shape_field(
            &peer,
            "page:1:shape:forged",
            "parentId",
            "page:1:shape:forged",
        );
        assert!(
            session
                .apply_update_v1(&peer_update(&session, &peer))
                .is_err()
        );
        assert_eq!(before, session.encode_state_as_update_v1());
    }

    #[test]
    fn remote_new_shapes_with_a_two_shape_parent_cycle_are_rejected() {
        let session = session();
        let before = session.encode_state_as_update_v1();
        let peer = peer_doc(&session, 9);
        write_peer_new_shape(&peer, "page:1:shape:forged-a", "page:1", "session", 1.0);
        write_peer_new_shape(&peer, "page:1:shape:forged-b", "page:1", "session", 1.0);
        write_peer_shape_field(
            &peer,
            "page:1:shape:forged-a",
            "parentId",
            "page:1:shape:forged-b",
        );
        write_peer_shape_field(
            &peer,
            "page:1:shape:forged-b",
            "parentId",
            "page:1:shape:forged-a",
        );
        assert!(
            session
                .apply_update_v1(&peer_update(&session, &peer))
                .is_err()
        );
        assert_eq!(before, session.encode_state_as_update_v1());
    }

    #[test]
    fn snapshot_terminates_on_a_cyclic_parent_chain_instead_of_overflowing() {
        let session = session();
        add_child_shape(&session, "page:1:shape:cycle-a", "page:1:shape:1");
        add_child_shape(&session, "page:1:shape:cycle-b", "page:1:shape:cycle-a");
        write_peer_shape_field(
            session.yrs_doc(),
            "page:1:shape:cycle-a",
            "parentId",
            "page:1:shape:cycle-b",
        );
        assert!(
            session
                .snapshot()
                .unwrap_err()
                .to_string()
                .contains("cyclic parent chain")
        );
    }

    #[test]
    fn remote_new_cell_on_a_locked_target_is_rejected() {
        let session = session();
        add_cell(&session, "LockWidth", Some("1"), None);
        let before = session.encode_state_as_update_v1();
        let peer = peer_doc(&session, 9);
        write_peer_new_cell(&peer, "page:1:shape:1", "Width", "5");
        assert!(
            session
                .apply_update_v1(&peer_update(&session, &peer))
                .is_err()
        );
        assert_eq!(before, session.encode_state_as_update_v1());
    }

    #[test]
    fn remote_new_cell_carrying_a_guard_formula_is_rejected() {
        let session = session();
        let before = session.encode_state_as_update_v1();
        let peer = peer_doc(&session, 9);
        write_peer_new_cell(&peer, "page:1:shape:1", "Height", "GUARD(1)");
        assert!(
            session
                .apply_update_v1(&peer_update(&session, &peer))
                .is_err()
        );
        assert_eq!(before, session.encode_state_as_update_v1());
    }

    fn remove_peer_cell(peer: &Doc, shape_id: &str, cell_key: &str) {
        let mut txn = peer.transact_mut();
        let cells = shape_cells(&txn, shape_id);
        cells.remove(&mut txn, cell_key);
    }

    fn remove_peer_cell_field(peer: &Doc, shape_id: &str, cell_key: &str, field: &str) {
        let mut txn = peer.transact_mut();
        let cells = shape_cells(&txn, shape_id);
        let cell = match cells.get(&txn, cell_key) {
            Some(yrs::Out::YMap(cell)) => cell,
            _ => unreachable!(),
        };
        cell.remove(&mut txn, field);
    }

    fn delete_peer_shape(peer: &Doc, page_id: &str, shape_id: &str) {
        let mut txn = peer.transact_mut();
        let sheets = txn.get_map(SHEETS).unwrap();
        sheets.remove(&mut txn, shape_id);
        let pages = txn.get_map(PAGES).unwrap();
        let page = match pages.get(&txn, page_id) {
            Some(yrs::Out::YMap(page)) => page,
            _ => unreachable!(),
        };
        let shapes = match page.get(&txn, "shapes") {
            Some(yrs::Out::YArray(shapes)) => shapes,
            _ => unreachable!(),
        };
        let mut index = None;
        for candidate in 0..shapes.len(&txn) {
            if let Some(yrs::Out::Any(yrs::Any::String(value))) = shapes.get(&txn, candidate)
                && value.as_ref() == shape_id
            {
                index = Some(candidate);
                break;
            }
        }
        if let Some(index) = index {
            shapes.remove_range(&mut txn, index, 1);
        }
    }

    #[test]
    fn remote_delete_of_a_guarded_cell_is_rejected_while_its_shape_survives() {
        let session = session();
        add_cell(&session, "Width", Some("GUARD(1)"), None);
        let before = session.encode_state_as_update_v1();
        let peer = peer_doc(&session, 9);
        remove_peer_cell(&peer, "page:1:shape:1", "Width");
        assert!(
            session
                .apply_update_v1(&peer_update(&session, &peer))
                .is_err()
        );
        assert_eq!(before, session.encode_state_as_update_v1());
    }

    #[test]
    fn remote_delete_of_a_locked_cell_is_rejected_while_its_shape_survives() {
        let session = session();
        add_cell(&session, "LockWidth", Some("1"), None);
        add_cell(&session, "Width", Some("5"), None);
        let before = session.encode_state_as_update_v1();
        let peer = peer_doc(&session, 9);
        remove_peer_cell(&peer, "page:1:shape:1", "Width");
        assert!(
            session
                .apply_update_v1(&peer_update(&session, &peer))
                .is_err()
        );
        assert_eq!(before, session.encode_state_as_update_v1());
    }

    #[test]
    fn remote_delete_of_a_baseline_formula_is_rejected_while_its_shape_survives() {
        let session = session();
        add_shape_cell(
            &session,
            "page:1:shape:1",
            "Width",
            None,
            None,
            Some("2"),
            None,
        );
        let before = session.encode_state_as_update_v1();
        let peer = peer_doc(&session, 9);
        remove_peer_cell_field(&peer, "page:1:shape:1", "Width", "baselineFormula");
        assert!(
            session
                .apply_update_v1(&peer_update(&session, &peer))
                .is_err()
        );
        assert_eq!(before, session.encode_state_as_update_v1());
    }

    #[test]
    fn remote_deletion_of_a_locked_shape_is_rejected() {
        let session = session();
        add_cell(&session, "LockDelete", Some("1"), None);
        let before = session.encode_state_as_update_v1();
        let peer = peer_doc(&session, 9);
        delete_peer_shape(&peer, "page:1", "page:1:shape:1");
        assert!(
            session
                .apply_update_v1(&peer_update(&session, &peer))
                .is_err()
        );
        assert_eq!(before, session.encode_state_as_update_v1());
    }

    /// Rejects remote changes that silently remove a protected cell.
    #[test]
    fn remote_update_that_disables_a_lock_forgets_the_cell_it_was_protecting() {
        let session = session();
        add_cell(&session, "LockWidth", Some("1"), None);
        add_cell(&session, "Width", Some("5"), None);
        let before = session.encode_state_as_update_v1();
        let peer = peer_doc(&session, 9);
        write_peer_cell_field(&peer, "page:1:shape:1", "LockWidth", "formula", "0");
        assert!(
            session
                .apply_update_v1(&peer_update(&session, &peer))
                .is_err()
        );
        assert_eq!(before, session.encode_state_as_update_v1());
    }

    /// Matches local deletion policy for unrelated guarded cells.
    #[test]
    fn remote_deletion_of_an_unlocked_shape_with_a_guarded_cell_is_accepted() {
        let session = session();
        add_cell(&session, "Width", Some("GUARD(1)"), None);
        let peer = peer_doc(&session, 9);
        delete_peer_shape(&peer, "page:1", "page:1:shape:1");
        assert!(
            session
                .apply_update_v1(&peer_update(&session, &peer))
                .is_ok()
        );
        assert!(
            session
                .snapshot()
                .unwrap()
                .pages
                .iter()
                .find(|page| page.id == "page:1")
                .unwrap()
                .shapes
                .iter()
                .all(|shape| shape.id != "page:1:shape:1")
        );
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
    fn concurrent_added_shapes_converge_without_identity_collisions() {
        let seed = session();
        let state = seed.encode_state_as_update_v1();
        let left = DiagramSession::open_from_update(&state, 11).unwrap();
        let right = DiagramSession::open_from_update(&state, 12).unwrap();
        let draft = ShapeDraft {
            name: Some("Added".to_owned()),
            cells: Vec::new(),
        };
        let left_added = left
            .add_shape(&EditCtx::local("left"), "page:1", &draft)
            .unwrap();
        let right_added = right
            .add_shape(&EditCtx::local("right"), "page:1", &draft)
            .unwrap();
        assert_ne!(left_added.shape_id, right_added.shape_id);
        let left_update = left
            .encode_diff_v1(&right.encode_state_vector_v1())
            .unwrap();
        let right_update = right
            .encode_diff_v1(&left.encode_state_vector_v1())
            .unwrap();
        left.apply_update_v1(&right_update).unwrap();
        right.apply_update_v1(&left_update).unwrap();
        assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
        assert_eq!(left.snapshot().unwrap().pages[0].shapes.len(), 4);
    }

    #[test]
    fn peers_converge_after_editing_distinct_section_row_cells() {
        let seed = session();
        let x = CellLocator {
            sheet: CellSheet::Page(1),
            shape_id: Some(1),
            section: Some("Geometry".to_owned()),
            section_index: None,
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
    fn group_subshape_snapshot_and_story_match_render_resolution() {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/group-master-shape.vsdx");
        let package = vsdx_parse::parse_vsdx(source).unwrap();
        let page = &package.page_part_paths[0];
        let resolver = vsdx_resolve::Resolver::new(&package);
        let resolved = resolver.resolve_page_shapes(page).unwrap();
        let session = DiagramSession::open(source, 7).unwrap();
        let snapshot = session.snapshot().unwrap();
        let group = &snapshot.pages[0].shapes[0];
        let source_group = package.page_contents[page].shapes().next().unwrap();
        let mut source_children = source_group.shapes();
        let direct = source_children.next().unwrap();
        let nested = source_children.next().unwrap().shapes().next().unwrap();
        let txn = session.doc.transact();
        let stories = txn.get_map(STORIES).unwrap();
        for (child, shape) in [
            (&group.children[0], direct),
            (&group.children[1].children[0], nested),
        ] {
            for (name, expected) in [("PinX", "1"), ("Width", "2"), ("PageValue", "23")] {
                let cell = child
                    .cells
                    .iter()
                    .find(|cell| cell.name == name && cell.locator.section.is_none())
                    .unwrap();
                assert_eq!(cell.value.as_deref(), Some(expected));
                let vsdx_resolve::Lookup::Found(render_cell) =
                    &resolved[&child.source_id].cells[name]
                else {
                    panic!("missing {name}")
                };
                assert_eq!(cell.value, render_cell.cell.value);
            }
            let Some(yrs::Out::Any(Any::String(story))) = stories.get(&txn, &child.id) else {
                panic!("missing story")
            };
            assert_eq!(story.as_ref(), "group label");
            let tokens = resolver
                .resolve_text_in_context(
                    shape,
                    &package.page_contents[page],
                    &resolved[&child.source_id],
                )
                .unwrap();
            let vsdx_resolve::ResolvedTextToken::CharacterRun { properties, .. } = &tokens[0]
            else {
                panic!("missing character run")
            };
            let vsdx_resolve::Lookup::Found(size) = &properties["Size"] else {
                panic!("missing font size")
            };
            assert_eq!(size.cell.value.as_deref(), Some("0.25"));
            assert_eq!(size.provenance, vsdx_resolve::Provenance::Page);
            let snapshot_size = child
                .cells
                .iter()
                .find(|cell| {
                    cell.name == "Size" && cell.locator.section.as_deref() == Some("Character")
                })
                .unwrap();
            assert_eq!(snapshot_size.value, size.cell.value);
            assert_eq!(
                tokens[1],
                vsdx_resolve::ResolvedTextToken::Literal("group label".into())
            );
        }
        let renderer = vsdx_render::Renderer::default();
        assert_eq!(
            renderer.layout_page(&package, page).unwrap(),
            renderer
                .layout_page(&session.package().unwrap(), page)
                .unwrap()
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
    fn deleting_a_groups_last_child_materializes_an_empty_shapes_collection() {
        let session = DiagramSession::open(
            include_bytes!("../../vsdx-parse/tests/fixtures/nested-groups.vsdx"),
            7,
        )
        .unwrap();
        let parent_id = session.snapshot().unwrap().pages[0].shapes[0].id.clone();
        let mut txn = session.doc.transact_mut();
        let sheets = txn.get_map(SHEETS).unwrap();
        let parent = match sheets.get(&txn, &parent_id) {
            Some(yrs::Out::YMap(parent)) => parent,
            _ => unreachable!(),
        };
        let children = match parent.get(&txn, "shapes") {
            Some(yrs::Out::YArray(children)) => children,
            _ => unreachable!(),
        };
        let child_count = children.len(&txn);
        children.remove_range(&mut txn, 0, child_count);
        drop(txn);
        assert!(
            session.snapshot().unwrap().pages[0].shapes[0]
                .children
                .is_empty()
        );
        let package = session.package().unwrap();
        let sheet = package.page_contents.get("visio/pages/page1.xml").unwrap();
        assert_eq!(sheet.shapes().next().unwrap().shapes().count(), 0);
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

    #[test]
    fn seeded_serializable_documents_reopen_to_the_live_projection() {
        let mut state = 0x5eed_cafe_u64;
        for case in 0..32_u64 {
            let source = match case % 4 {
                0 => include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx").as_slice(),
                1 => {
                    include_bytes!("../../vsdx-parse/tests/fixtures/nested-groups.vsdx").as_slice()
                }
                2 => include_bytes!("../../../apps/demo/public/betteroffice-demo.vsdx").as_slice(),
                _ => include_bytes!("../../vsdx-parse/tests/fixtures/grouped-glue.vsdx").as_slice(),
            };
            let session = DiagramSession::open(source, 100 + case).unwrap();
            let edits = 1 + next_test_random(&mut state) % 8;
            for _ in 0..edits {
                apply_generated_edit(&session, &mut state, case % 4 == 0);
            }
            let live = session.snapshot().unwrap();
            let update = session.encode_state_as_update_v1();
            let validated = DiagramSession::open_from_update(&update, 1_000 + case).unwrap();
            assert_eq!(validated.snapshot().unwrap(), live, "case {case}");
            let saved = session.save().unwrap();
            let reopened = DiagramSession::open(&saved, 10_000 + case).unwrap();
            assert_reopened_projection_eq(&session, &reopened, "case {case}");
            let left = DiagramSession::open_from_update(&update, 20_000 + case).unwrap();
            let right = DiagramSession::open_from_update(&update, 30_000 + case).unwrap();
            add_generated_shape(&left, &mut state);
            add_generated_shape(&right, &mut state);
            let left_update = left
                .encode_diff_v1(&right.encode_state_vector_v1())
                .unwrap();
            let right_update = right
                .encode_diff_v1(&left.encode_state_vector_v1())
                .unwrap();
            left.apply_update_v1(&right_update).unwrap();
            right.apply_update_v1(&left_update).unwrap();
            assert_eq!(
                left.snapshot().unwrap(),
                right.snapshot().unwrap(),
                "case {case}"
            );
            let saved = left.save().unwrap();
            let reopened = DiagramSession::open(&saved, 40_000 + case).unwrap();
            assert_reopened_projection_eq(&left, &reopened, "merged case {case}");
        }
    }

    fn assert_reopened_projection_eq(live: &DiagramSession, reopened: &DiagramSession, case: &str) {
        assert_semantic_snapshot_eq(
            &live.snapshot().unwrap(),
            &reopened.snapshot().unwrap(),
            case,
        );
        let live_package = live.package().unwrap();
        let reopened_package = reopened.package().unwrap();
        let renderer = vsdx_render::Renderer::default();
        assert_eq!(
            live_package.page_part_paths.len(),
            reopened_package.page_part_paths.len()
        );
        for (page_index, (live_path, reopened_path)) in live_package
            .page_part_paths
            .iter()
            .zip(&reopened_package.page_part_paths)
            .enumerate()
        {
            assert_eq!(
                renderer.layout_page(&live_package, live_path).unwrap(),
                renderer
                    .layout_page(&reopened_package, reopened_path)
                    .unwrap(),
                "{case}, page {page_index}"
            );
        }
    }

    fn assert_semantic_snapshot_eq(live: &DiagramSnapshot, reopened: &DiagramSnapshot, case: &str) {
        assert_eq!(live.pages.len(), reopened.pages.len(), "{case}");
        for (live_page, reopened_page) in live.pages.iter().zip(&reopened.pages) {
            assert_eq!(
                live_page.source_part_path, reopened_page.source_part_path,
                "{case}"
            );
            assert_eq!(live_page.name, reopened_page.name, "{case}");
            assert_semantic_shapes_eq(&live_page.shapes, &reopened_page.shapes, case);
        }
    }

    fn assert_semantic_shapes_eq(live: &[ShapeSnapshot], reopened: &[ShapeSnapshot], case: &str) {
        assert_eq!(live.len(), reopened.len(), "{case}");
        for (live_shape, reopened_shape) in live.iter().zip(reopened) {
            assert_eq!(live_shape.source_id, reopened_shape.source_id, "{case}");
            assert_eq!(live_shape.name, reopened_shape.name, "{case}");
            assert_eq!(live_shape.cells, reopened_shape.cells, "{case}");
            assert_semantic_shapes_eq(&live_shape.children, &reopened_shape.children, case);
        }
    }

    #[test]
    fn deleting_a_group_removes_connects_from_live_and_saved_projections() {
        let session = DiagramSession::open(
            include_bytes!("../../vsdx-parse/tests/fixtures/grouped-glue.vsdx"),
            601,
        )
        .unwrap();
        session
            .delete_shape(&EditCtx::local("local"), "page:1", "page:1:shape:10")
            .unwrap();
        let saved = session.save().unwrap();
        let reopened = DiagramSession::open(&saved, 602).unwrap();
        assert_reopened_projection_eq(&session, &reopened, "local group deletion");
        assert_no_connects_to_deleted_group(&session);
    }

    #[test]
    fn accepted_peer_group_deletion_removes_connects_from_live_and_saved_projections() {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/grouped-glue.vsdx");
        let local = DiagramSession::open(source, 602).unwrap();
        let peer = DiagramSession::open(source, 603).unwrap();
        peer.delete_shape(&EditCtx::local("peer"), "page:1", "page:1:shape:10")
            .unwrap();
        local
            .apply_update_v1(
                &peer
                    .encode_diff_v1(&local.encode_state_vector_v1())
                    .unwrap(),
            )
            .unwrap();
        let saved = local.save().unwrap();
        let reopened = DiagramSession::open(&saved, 604).unwrap();
        assert_reopened_projection_eq(&local, &reopened, "peer group deletion");
        assert_no_connects_to_deleted_group(&local);
    }

    fn assert_no_connects_to_deleted_group(session: &DiagramSession) {
        assert!(
            session.package().unwrap().page_contents["visio/pages/page1.xml"]
                .connects()
                .all(|connect| ![10, 11].contains(&connect.from_sheet)
                    && ![10, 11].contains(&connect.to_sheet))
        );
    }

    fn next_test_random(state: &mut u64) -> u64 {
        *state = state
            .wrapping_mul(6_364_136_223_846_793_005)
            .wrapping_add(1);
        *state >> 32
    }

    fn apply_generated_edit(session: &DiagramSession, state: &mut u64, allow_addition: bool) {
        let snapshot = session.snapshot().unwrap();
        let page = &snapshot.pages[(next_test_random(state) as usize) % snapshot.pages.len()];
        let shapes = shape_choices(&page.shapes);
        let context = EditCtx::local("seeded");
        match next_test_random(state) % 5 {
            0 => {
                let formula = (1 + next_test_random(state) % 10_000).to_string();
                if let Some((shape, cell)) = shapes
                    .iter()
                    .find_map(|(shape, _)| shape.cells.first().map(|cell| (shape, cell)))
                {
                    session
                        .set_cell_formula_at(
                            &context,
                            &page.id,
                            &shape.id,
                            cell.locator.clone(),
                            formula,
                        )
                        .unwrap();
                }
            }
            1 => {
                if !allow_addition {
                    return;
                }
                session
                    .add_shape(
                        &context,
                        &page.id,
                        &ShapeDraft {
                            name: Some("Generated".to_owned()),
                            cells: generated_shape_cells(),
                        },
                    )
                    .unwrap();
            }
            2 => {
                if let Some((shape, _)) =
                    shapes.get((next_test_random(state) as usize) % shapes.len().max(1))
                {
                    session.delete_shape(&context, &page.id, &shape.id).unwrap();
                }
            }
            3 => {
                if let Some((shape, sibling_len)) =
                    shapes.get((next_test_random(state) as usize) % shapes.len().max(1))
                {
                    session
                        .reorder_shape(
                            &context,
                            &page.id,
                            &shape.id,
                            sibling_len.saturating_sub(1) as u32,
                        )
                        .unwrap();
                }
            }
            _ => {
                if snapshot.pages.len() > 1 {
                    session.reorder_page(&context, &page.id, 0).unwrap();
                }
            }
        }
    }

    fn shape_choices(shapes: &[ShapeSnapshot]) -> Vec<(&ShapeSnapshot, usize)> {
        let mut result = Vec::new();
        let mut pending = vec![shapes];
        while let Some(siblings) = pending.pop() {
            for shape in siblings {
                result.push((shape, siblings.len()));
                if !shape.children.is_empty() {
                    pending.push(&shape.children);
                }
            }
        }
        result
    }

    fn add_generated_shape(session: &DiagramSession, state: &mut u64) {
        let snapshot = session.snapshot().unwrap();
        let page = &snapshot.pages[(next_test_random(state) as usize) % snapshot.pages.len()];
        session
            .add_shape(
                &EditCtx::local("seeded"),
                &page.id,
                &ShapeDraft {
                    name: Some("Generated".to_owned()),
                    cells: generated_shape_cells(),
                },
            )
            .unwrap();
    }

    fn generated_shape_cells() -> Vec<CellSnapshot> {
        [
            ("LocPinX", "Width * 0.5"),
            ("LocPinY", "Height * 0.5"),
            ("PageHeight", "11"),
            ("PageWidth", "8.5"),
            ("TxtPinX", "Width * 0.5"),
            ("TxtPinY", "Height * 0.5"),
        ]
        .into_iter()
        .map(|(name, formula)| CellSnapshot {
            row_type: None,
            locator: CellLocator {
                sheet: CellSheet::Page(1),
                shape_id: None,
                section: None,
                section_index: None,
                row: None,
                cell_name: name.to_owned(),
            },
            name: name.to_owned(),
            formula: Some(formula.to_owned()),
            value: None,
        })
        .collect()
    }
}
