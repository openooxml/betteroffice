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
    #[cfg(test)]
    pub(crate) fn undo_depth(&self) -> usize {
        self.undo.borrow().undo_depth()
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
        add_cell_to(session, "page:1:shape:1", name, formula, value);
    }

    fn add_cell_to(
        session: &DiagramSession,
        shape_id: &str,
        name: &str,
        formula: Option<&str>,
        value: Option<&str>,
    ) {
        add_cell_at_to(session, shape_id, name, None, None, formula, value);
    }

    fn add_cell_at(
        session: &DiagramSession,
        name: &str,
        section: Option<&str>,
        row: Option<CellRow>,
        formula: Option<&str>,
        value: Option<&str>,
    ) {
        add_cell_at_to(
            session,
            "page:1:shape:1",
            name,
            section,
            row,
            formula,
            value,
        );
    }

    fn add_cell_at_to(
        session: &DiagramSession,
        shape_id: &str,
        name: &str,
        section: Option<&str>,
        row: Option<CellRow>,
        formula: Option<&str>,
        value: Option<&str>,
    ) {
        let mut txn = session.doc.transact_mut_with(HYDRATE_ORIGIN);
        let sheets = txn.get_map(SHEETS).unwrap();
        let shape = match sheets.get(&txn, shape_id) {
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
    fn identical_formula_rewrites_skip_the_undo_stack() {
        let session = session();
        add_shape_cell(
            &session,
            "page:1:shape:1",
            "PinX",
            None,
            None,
            Some("3"),
            Some("3"),
        );
        let receipt = session
            .set_cell_formula(
                &EditCtx::local("test"),
                "page:1",
                "page:1:shape:1",
                "PinX",
                "3",
            )
            .unwrap();
        assert_eq!(receipt.before.as_deref(), Some("3"));
        assert!(!session.can_undo());
        session
            .set_cell_formula(
                &EditCtx::local("test"),
                "page:1",
                "page:1:shape:1",
                "PinX",
                "4",
            )
            .unwrap();
        assert!(session.can_undo());
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
                        &ShapeDraft {
                            name: None,
                            master: None,
                            cells
                        }
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
            master: None,
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

    const STENCIL_SOURCE: &[u8] =
        include_bytes!("../../vsdx-parse/tests/fixtures/document-stencil.vsdx");

    fn stencil_session() -> DiagramSession {
        DiagramSession::open(STENCIL_SOURCE, 11).unwrap()
    }

    fn placement_cell(name: &str, formula: &str) -> CellSnapshot {
        CellSnapshot {
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
        }
    }

    fn stencil_instance_draft(master: Option<u32>) -> ShapeDraft {
        ShapeDraft {
            name: Some("Stencil instance".to_owned()),
            master,
            cells: [
                ("PinX", "4"),
                ("PinY", "5"),
                ("Width", "2"),
                ("Height", "1"),
                ("LocPinX", "1"),
                ("LocPinY", "0.5"),
            ]
            .into_iter()
            .map(|(name, formula)| placement_cell(name, formula))
            .collect(),
        }
    }

    fn shape_paths(list: &vsdx_render::VsdxDisplayList, id: &str) -> Vec<String> {
        list.primitives
            .iter()
            .filter_map(|primitive| match primitive {
                vsdx_render::Primitive::Shape {
                    id: shape, path, ..
                } if shape.as_str() == id => Some(format!("{path:?}")),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn document_stencil_lists_master_names() {
        let names = stencil_session().package().unwrap().master_names;
        assert_eq!(
            names.into_iter().collect::<Vec<_>>(),
            [
                (1, "Stencil-Rect".to_owned()),
                (2, "Stencil-Tri".to_owned())
            ]
        );
    }

    #[test]
    fn master_instance_insert_resolves_geometry_through_the_master() {
        let session = stencil_session();
        let receipt = session
            .add_shape(
                &EditCtx::local("test"),
                "page:1",
                &stencil_instance_draft(Some(1)),
            )
            .unwrap();
        let snapshot = session.snapshot().unwrap();
        let added = snapshot.pages[0]
            .shapes
            .iter()
            .find(|shape| shape.id == receipt.shape_id)
            .unwrap();
        assert_eq!(added.master, Some(1));
        assert!(
            added
                .cells
                .iter()
                .all(|cell| cell.locator.section.as_deref() != Some("Geometry")),
            "a master instance stores placement only; geometry stays inherited: {:?}",
            added
                .cells
                .iter()
                .map(|cell| &cell.name)
                .collect::<Vec<_>>()
        );
        let renderer = vsdx_render::Renderer::default();
        let package = session.package().unwrap();
        let list = renderer
            .layout_page(&package, "visio/pages/page1.xml")
            .unwrap();
        let inserted = shape_paths(&list, &format!("visio/pages/page1.xml:{}", added.source_id));
        assert_eq!(
            inserted.len(),
            1,
            "the instance renders its inherited geometry"
        );
        assert_eq!(
            inserted,
            shape_paths(&list, "visio/pages/page1.xml:1"),
            "same master plus same placement renders the same path"
        );
    }

    #[test]
    fn master_instance_insert_rejects_unknown_masters() {
        let session = stencil_session();
        let before = session.snapshot().unwrap();
        for master in [Some(0), Some(999)] {
            assert!(
                session
                    .add_shape(
                        &EditCtx::local("test"),
                        "page:1",
                        &stencil_instance_draft(master)
                    )
                    .is_err(),
                "master {master:?} must be refused"
            );
        }
        assert_eq!(session.snapshot().unwrap(), before);
    }

    #[test]
    fn master_instance_insert_round_trips_through_save() {
        let session = stencil_session();
        let receipt = session
            .add_shape(
                &EditCtx::local("test"),
                "page:1",
                &stencil_instance_draft(Some(2)),
            )
            .unwrap();
        let saved = session.save().unwrap();
        let reparsed = vsdx_parse::parse_vsdx(&saved).unwrap();
        let part = reparsed.page_part_paths[0].clone();
        let stored = reparsed.page_contents[&part]
            .shapes()
            .max_by_key(|shape| shape.id)
            .unwrap();
        assert_eq!(stored.master, Some(2));
        let reopened = DiagramSession::open(&saved, 12).unwrap();
        let reopened_snapshot = reopened.snapshot().unwrap();
        let live_source = session.snapshot().unwrap().pages[0]
            .shapes
            .iter()
            .find(|shape| shape.id == receipt.shape_id)
            .unwrap()
            .source_id;
        let added = reopened_snapshot.pages[0]
            .shapes
            .iter()
            .find(|shape| shape.source_id == live_source)
            .unwrap();
        assert_eq!(added.master, Some(2));
        let renderer = vsdx_render::Renderer::default();
        let live = renderer
            .layout_page(&session.package().unwrap(), "visio/pages/page1.xml")
            .unwrap();
        let again = renderer
            .layout_page(&reopened.package().unwrap(), "visio/pages/page1.xml")
            .unwrap();
        assert_eq!(live, again);
    }

    #[test]
    fn master_instance_insert_undoes_in_one_step() {
        let session = stencil_session();
        let before = session.snapshot().unwrap();
        session
            .add_shape(
                &EditCtx::local("test"),
                "page:1",
                &stencil_instance_draft(Some(1)),
            )
            .unwrap();
        assert!(session.undo());
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(!session.can_undo());
    }

    fn placement_only_draft(master: u32, width: Option<&str>) -> ShapeDraft {
        let mut cells = vec![placement_cell("PinX", "4"), placement_cell("PinY", "2")];
        if let Some(width) = width {
            cells.push(placement_cell("Width", width));
            cells.push(placement_cell("Height", width));
            cells.push(placement_cell("LocPinX", "Width*0.5"));
            cells.push(placement_cell("LocPinY", "Height*0.5"));
        }
        ShapeDraft {
            name: None,
            master: Some(master),
            cells,
        }
    }

    #[test]
    fn master_instance_insert_keeps_the_master_dimensions_when_none_are_written() {
        let session = stencil_session();
        let inherited = session
            .add_shape(
                &EditCtx::local("test"),
                "page:1",
                &placement_only_draft(2, None),
            )
            .unwrap();
        let squared = session
            .add_shape(
                &EditCtx::local("test"),
                "page:1",
                &placement_only_draft(2, Some("1")),
            )
            .unwrap();
        let snapshot = session.snapshot().unwrap();
        let source_id = |shape_id: &str| {
            snapshot.pages[0]
                .shapes
                .iter()
                .find(|shape| shape.id == shape_id)
                .unwrap()
                .source_id
        };
        let list = vsdx_render::Renderer::default()
            .layout_page(&session.package().unwrap(), "visio/pages/page1.xml")
            .unwrap();
        let inherited = shape_paths(
            &list,
            &format!("visio/pages/page1.xml:{}", source_id(&inherited.shape_id)),
        );
        assert_eq!(
            inherited,
            shape_paths(&list, "visio/pages/page1.xml:3"),
            "an instance without an XForm renders like the stored instance of the same master"
        );
        assert_ne!(
            inherited,
            shape_paths(
                &list,
                &format!("visio/pages/page1.xml:{}", source_id(&squared.shape_id))
            ),
            "writing 1x1 distorts a non-square master"
        );
    }

    fn write_peer_shape_master(peer: &Doc, shape_id: &str, master: f64) {
        let mut txn = peer.transact_mut();
        let sheets = txn.get_map(SHEETS).unwrap();
        let shape = match sheets.get(&txn, shape_id) {
            Some(yrs::Out::YMap(shape)) => shape,
            _ => unreachable!(),
        };
        shape.insert(&mut txn, "master", master);
    }

    #[test]
    fn remote_master_instances_are_validated_against_the_package() {
        for (master, accepted) in [(1.0, true), (0.0, false), (999.0, false)] {
            let session = stencil_session();
            let before = session.encode_state_as_update_v1();
            let peer =
                DiagramSession::open_from_update(&session.encode_state_as_update_v1(), 8).unwrap();
            let receipt = peer
                .add_shape(
                    &EditCtx::local("peer"),
                    "page:1",
                    &stencil_instance_draft(Some(1)),
                )
                .unwrap();
            if master != 1.0 {
                write_peer_shape_master(&peer.doc, &receipt.shape_id, master);
            }
            let update = peer
                .encode_diff_v1(&session.encode_state_vector_v1())
                .unwrap();
            assert_eq!(
                session.apply_update_v1(&update).is_ok(),
                accepted,
                "master {master}"
            );
            if !accepted {
                assert_eq!(before, session.encode_state_as_update_v1());
            }
        }
    }

    #[test]
    fn remote_rewrite_of_an_existing_shape_master_is_rejected() {
        let session = stencil_session();
        let before = session.encode_state_as_update_v1();
        let peer = peer_doc(&session, 9);
        write_peer_shape_master(&peer, "page:1:shape:1", 2.0);
        assert!(
            session
                .apply_update_v1(&peer_update(&session, &peer))
                .is_err()
        );
        assert_eq!(before, session.encode_state_as_update_v1());
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
    fn converging_redirects_refuse_the_batch_without_writing() {
        for (source, target) in [("PinX", "Width"), ("Width", "Height")] {
            let session = session();
            for cell in ["PinX", "PinY", "Width", "Height"] {
                add_cell(&session, cell, Some("1"), None);
            }
            add_cell(&session, source, Some(&format!("SETATREF({target})")), None);
            let before = session.snapshot().unwrap();
            let vector = session.encode_state_vector_v1();
            assert_eq!(
                session
                    .set_shape_bounds(
                        &EditCtx::local("a"),
                        "page:1",
                        "page:1:shape:1",
                        ["2", "3", "4", "5"].map(str::to_owned)
                    )
                    .unwrap_err()
                    .to_string(),
                format!("invalid diagram state: redirects converge on {target} more than once")
            );
            assert_eq!(session.snapshot().unwrap(), before);
            assert_eq!(session.encode_state_vector_v1(), vector);
        }
    }

    #[test]
    fn shape_bounds_stop_at_a_missing_cell_without_writing_the_others() {
        let session = session();
        for cell in ["PinX", "PinY", "Width"] {
            add_cell(&session, cell, Some("1"), None);
        }
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
    fn resize_loc_pin_refuses_formulas_outside_the_shape_sheet() {
        for formula in ["ThePage!PageWidth*0.5", "User.Anchor"] {
            let session = session();
            for (name, value) in [("Width", "2"), ("Height", "3")] {
                add_cell(&session, name, Some(value), None);
            }
            add_cell(&session, "LocPinX", Some(formula), Some("1"));
            assert_eq!(
                session
                    .resize_loc_pin("page:1", "page:1:shape:1", 4.0, 6.0)
                    .unwrap_err()
                    .to_string(),
                "invalid diagram state: cannot evaluate LocPinX for resize",
                "{formula}"
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
    fn lock_rotate_refuses_angle_writes() {
        let session = session();
        add_cell(&session, "Angle", Some("0"), None);
        add_cell(&session, "LockRotate", Some("1"), None);
        let error = session
            .set_cell_formula(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:1",
                "Angle",
                "1",
            )
            .unwrap_err();
        assert!(
            error
                .to_string()
                .contains("LockRotate protects this rotate gesture")
        );
        let cells = &session.snapshot().unwrap().pages[0].shapes[0].cells;
        assert_eq!(
            cells
                .iter()
                .find(|cell| cell.name == "Angle")
                .unwrap()
                .formula
                .as_deref(),
            Some("0")
        );
    }

    #[test]
    fn angle_writes_apply_without_lock_rotate() {
        let session = session();
        add_cell(&session, "Angle", Some("0"), None);
        let receipt = session
            .set_cell_formula(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:1",
                "Angle",
                "1",
            )
            .unwrap();
        assert_eq!(receipt.cell_name, "Angle");
        let cells = &session.snapshot().unwrap().pages[0].shapes[0].cells;
        assert_eq!(
            cells
                .iter()
                .find(|cell| cell.name == "Angle")
                .unwrap()
                .formula
                .as_deref(),
            Some("1")
        );
    }

    fn loc_pin_at_size(session: &DiagramSession, width: f64, height: f64) -> (f64, f64) {
        let txn = session.doc.transact();
        crate::diagram::loc_pin_at_size(&txn, "page:1:shape:1", width, height).unwrap()
    }

    #[test]
    fn loc_pin_at_size_holds_when_pin_flows_through_an_intermediate_cell() {
        let indirect = |offset: &str| {
            let session = session();
            for (name, formula) in [
                ("Width", "2"),
                ("Height", "1"),
                ("PinX", "5"),
                ("LocPinX", "User.Offset*Width"),
                ("LocPinY", "0.5"),
            ] {
                add_cell(&session, name, Some(formula), None);
            }
            add_cell_at(
                &session,
                "Value",
                Some("User"),
                Some(CellRow::Name("Offset".to_owned())),
                Some(offset),
                None,
            );
            loc_pin_at_size(&session, 9.0, 1.0).0
        };
        assert_eq!(indirect("PinX*0.5"), 5.0);
        assert_eq!(indirect("0.25"), 2.25);
    }

    fn count_updates(
        session: &DiagramSession,
    ) -> (std::sync::Arc<std::sync::Mutex<usize>>, Subscription) {
        let updates = std::sync::Arc::new(std::sync::Mutex::new(0usize));
        let observed = std::sync::Arc::clone(&updates);
        let subscription = session
            .observe_update_v1(move |_| {
                *observed.lock().unwrap() += 1;
            })
            .unwrap();
        (updates, subscription)
    }

    #[test]
    fn set_shape_bounds_commits_size_and_pin_in_a_single_update() {
        let session = session();
        for name in ["Width", "Height", "PinX", "PinY"] {
            add_cell(&session, name, Some("1"), None);
        }
        let (updates, _subscription) = count_updates(&session);
        let receipts = session
            .set_shape_bounds(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:1",
                [
                    "4".to_owned(),
                    "5".to_owned(),
                    "2".to_owned(),
                    "3".to_owned(),
                ],
            )
            .unwrap();
        assert_eq!(
            receipts
                .iter()
                .map(|receipt| receipt.cell_name.as_str())
                .collect::<Vec<_>>(),
            ["PinX", "PinY", "Width", "Height"]
        );
        assert_eq!(*updates.lock().unwrap(), 1);
    }

    #[test]
    fn set_shape_bounds_undoes_size_and_pin_together() {
        let session = session();
        for name in ["Width", "Height", "PinX", "PinY"] {
            add_cell(&session, name, Some("1"), None);
        }
        session
            .set_shape_bounds(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:1",
                [
                    "4".to_owned(),
                    "5".to_owned(),
                    "2".to_owned(),
                    "3".to_owned(),
                ],
            )
            .unwrap();
        assert_eq!(session.undo_depth(), 1);
        assert!(session.undo());
        let cells = &session.snapshot().unwrap().pages[0].shapes[0].cells;
        for name in ["Width", "Height", "PinX", "PinY"] {
            assert_eq!(
                cells
                    .iter()
                    .find(|cell| cell.name == name)
                    .unwrap()
                    .formula
                    .as_deref(),
                Some("1")
            );
        }
        assert!(!session.can_undo());
    }

    #[test]
    fn move_shapes_undoes_every_pin_together() {
        let session = session();
        for shape_id in ["page:1:shape:1", "page:1:shape:2"] {
            add_cell_to(&session, shape_id, "PinX", Some("1"), None);
            add_cell_to(&session, shape_id, "PinY", Some("1"), None);
        }
        let (updates, _subscription) = count_updates(&session);
        session
            .move_shapes(
                &EditCtx::local("a"),
                &[
                    ShapeMove {
                        page_id: "page:1".to_owned(),
                        shape_id: "page:1:shape:1".to_owned(),
                        x: "2".to_owned(),
                        y: "3".to_owned(),
                    },
                    ShapeMove {
                        page_id: "page:1".to_owned(),
                        shape_id: "page:1:shape:2".to_owned(),
                        x: "4".to_owned(),
                        y: "5".to_owned(),
                    },
                ],
            )
            .unwrap();
        assert_eq!(*updates.lock().unwrap(), 1);
        assert_eq!(session.undo_depth(), 1);
        assert!(session.undo());
        let shapes = &session.snapshot().unwrap().pages[0].shapes;
        for shape in shapes {
            for name in ["PinX", "PinY"] {
                assert_eq!(
                    shape
                        .cells
                        .iter()
                        .find(|cell| cell.name == name)
                        .unwrap()
                        .formula
                        .as_deref(),
                    Some("1")
                );
            }
        }
        assert!(!session.can_undo());
    }

    #[test]
    fn delete_shapes_undoes_every_shape_together() {
        let session = session();
        for shape_id in ["page:1:shape:1", "page:1:shape:2"] {
            add_cell_to(&session, shape_id, "PinX", Some("1"), None);
        }
        let (updates, _subscription) = count_updates(&session);
        session
            .delete_shapes(
                &EditCtx::local("a"),
                &[
                    ShapeDelete {
                        page_id: "page:1".to_owned(),
                        shape_id: "page:1:shape:1".to_owned(),
                    },
                    ShapeDelete {
                        page_id: "page:1".to_owned(),
                        shape_id: "page:1:shape:2".to_owned(),
                    },
                ],
            )
            .unwrap();
        assert_eq!(*updates.lock().unwrap(), 1);
        assert_eq!(session.undo_depth(), 1);
        assert!(session.snapshot().unwrap().pages[0].shapes.is_empty());
        assert!(session.undo());
        assert_eq!(session.snapshot().unwrap().pages[0].shapes.len(), 2);
        assert!(!session.can_undo());
    }

    #[test]
    fn delete_shapes_drops_the_text_story_with_the_shape() {
        let session = session();
        add_cell_to(&session, "page:1:shape:1", "PinX", Some("1"), None);
        session
            .set_shape_text(&EditCtx::local("a"), "page:1", "page:1:shape:1", "label")
            .unwrap();
        session
            .delete_shapes(
                &EditCtx::local("a"),
                &[ShapeDelete {
                    page_id: "page:1".to_owned(),
                    shape_id: "page:1:shape:1".to_owned(),
                }],
            )
            .unwrap();
        {
            let txn = session.doc.transact();
            assert!(
                txn.get_map(STORIES)
                    .unwrap()
                    .get(&txn, "page:1:shape:1")
                    .is_none()
            );
        }
        let reopened =
            DiagramSession::open_from_update(&session.encode_state_as_update_v1(), 8).unwrap();
        let shapes = &reopened.snapshot().unwrap().pages[0].shapes;
        assert!(shapes.iter().all(|shape| shape.id != "page:1:shape:1"));
    }

    #[test]
    fn delete_shapes_with_a_missing_shape_leaves_the_document_untouched() {
        let session = session();
        add_cell_to(&session, "page:1:shape:1", "PinX", Some("1"), None);
        let before = session.snapshot().unwrap();
        let result = session.delete_shapes(
            &EditCtx::local("a"),
            &[
                ShapeDelete {
                    page_id: "page:1".to_owned(),
                    shape_id: "page:1:shape:1".to_owned(),
                },
                ShapeDelete {
                    page_id: "page:1".to_owned(),
                    shape_id: "missing".to_owned(),
                },
            ],
        );
        assert!(matches!(result, Err(EditError::ShapeNotFound(shape_id)) if shape_id == "missing"));
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(!session.can_undo());
    }

    #[test]
    fn move_shapes_refuses_the_whole_batch_when_one_shape_is_locked() {
        let session = session();
        for shape_id in ["page:1:shape:1", "page:1:shape:2"] {
            add_cell_to(&session, shape_id, "PinX", Some("1"), None);
            add_cell_to(&session, shape_id, "PinY", Some("1"), None);
        }
        add_cell_to(&session, "page:1:shape:2", "LockMoveX", Some("1"), None);
        let before = session.snapshot().unwrap();
        let (updates, _subscription) = count_updates(&session);
        let error = session
            .move_shapes(
                &EditCtx::local("a"),
                &[
                    ShapeMove {
                        page_id: "page:1".to_owned(),
                        shape_id: "page:1:shape:1".to_owned(),
                        x: "2".to_owned(),
                        y: "3".to_owned(),
                    },
                    ShapeMove {
                        page_id: "page:1".to_owned(),
                        shape_id: "page:1:shape:2".to_owned(),
                        x: "4".to_owned(),
                        y: "5".to_owned(),
                    },
                ],
            )
            .unwrap_err();
        assert!(error.to_string().contains("LockMoveX"));
        assert_eq!(*updates.lock().unwrap(), 0);
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(!session.can_undo());
    }

    #[test]
    fn delete_shapes_refuses_the_whole_batch_when_one_shape_is_locked() {
        let session = session();
        for shape_id in ["page:1:shape:1", "page:1:shape:2"] {
            add_cell_to(&session, shape_id, "PinX", Some("1"), None);
        }
        add_cell_to(&session, "page:1:shape:2", "LockDelete", Some("1"), None);
        let before = session.snapshot().unwrap();
        let (updates, _subscription) = count_updates(&session);
        let error = session
            .delete_shapes(
                &EditCtx::local("a"),
                &[
                    ShapeDelete {
                        page_id: "page:1".to_owned(),
                        shape_id: "page:1:shape:1".to_owned(),
                    },
                    ShapeDelete {
                        page_id: "page:1".to_owned(),
                        shape_id: "page:1:shape:2".to_owned(),
                    },
                ],
            )
            .unwrap_err();
        assert!(error.to_string().contains("LockDelete"));
        assert_eq!(*updates.lock().unwrap(), 0);
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(!session.can_undo());
    }

    #[test]
    fn set_cell_formulas_undoes_every_shape_together() {
        let session = session();
        for shape_id in ["page:1:shape:1", "page:1:shape:2"] {
            add_cell_to(&session, shape_id, "FillForegnd", Some("1"), None);
        }
        let (updates, _subscription) = count_updates(&session);
        session
            .set_cell_formulas(
                &EditCtx::local("a"),
                &[
                    CellFormulaWrite {
                        page_id: "page:1".to_owned(),
                        shape_id: "page:1:shape:1".to_owned(),
                        cell_name: "FillForegnd".to_owned(),
                        formula: "2".to_owned(),
                    },
                    CellFormulaWrite {
                        page_id: "page:1".to_owned(),
                        shape_id: "page:1:shape:2".to_owned(),
                        cell_name: "FillForegnd".to_owned(),
                        formula: "2".to_owned(),
                    },
                ],
            )
            .unwrap();
        assert_eq!(*updates.lock().unwrap(), 1);
        assert_eq!(session.undo_depth(), 1);
        assert!(session.undo());
        let shapes = &session.snapshot().unwrap().pages[0].shapes;
        for shape in shapes {
            assert_eq!(
                shape
                    .cells
                    .iter()
                    .find(|cell| cell.name == "FillForegnd")
                    .unwrap()
                    .formula
                    .as_deref(),
                Some("1")
            );
        }
        assert!(!session.can_undo());
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
    fn property_value_edits_write_through_the_edit_session() {
        let session = session();
        let row = CellRow::Name("Device".to_owned());
        add_cell_at(
            &session,
            "Label",
            Some("Property"),
            Some(row.clone()),
            None,
            Some("Device name"),
        );
        add_cell_at(
            &session,
            "Value",
            Some("Property"),
            Some(row.clone()),
            Some("\"Old\""),
            Some("Old"),
        );
        let locator = CellLocator {
            sheet: CellSheet::Page(1),
            shape_id: Some(1),
            section: Some("Property".to_owned()),
            section_index: None,
            row: Some(row),
            cell_name: "Value".to_owned(),
        };
        let receipt = session
            .set_cell_formula_at(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:1",
                locator.clone(),
                "\"New\"",
            )
            .unwrap();
        assert_eq!(receipt.cell_name, "Value");
        let snapshot = session.snapshot().unwrap();
        let cells = &snapshot.pages[0].shapes[0].cells;
        assert_eq!(
            cells
                .iter()
                .find(|cell| cell.locator == locator)
                .unwrap()
                .formula
                .as_deref(),
            Some("\"New\"")
        );
        assert_eq!(
            cells
                .iter()
                .find(|cell| cell.name == "Label"
                    && cell.locator.section.as_deref() == Some("Property"))
                .unwrap()
                .value
                .as_deref(),
            Some("Device name")
        );
    }

    fn property_row(
        session: &DiagramSession,
        name: &str,
        type_code: Option<&str>,
        formula: &str,
        value: &str,
    ) -> CellRow {
        let row = CellRow::Name(name.to_owned());
        add_cell_at(
            session,
            "Label",
            Some("Property"),
            Some(row.clone()),
            None,
            Some(name),
        );
        if let Some(code) = type_code {
            add_cell_at(
                session,
                "Type",
                Some("Property"),
                Some(row.clone()),
                None,
                Some(code),
            );
        }
        add_cell_at(
            session,
            "Value",
            Some("Property"),
            Some(row.clone()),
            Some(formula),
            Some(value),
        );
        row
    }

    fn data_write(row: &CellRow, formula: &str) -> ShapeDataWrite {
        ShapeDataWrite {
            row: row.clone(),
            section_index: None,
            formula: formula.to_owned(),
        }
    }

    fn value_formula(session: &DiagramSession, row: &CellRow) -> Option<String> {
        let snapshot = session.snapshot().unwrap();
        snapshot.pages[0].shapes[0]
            .cells
            .iter()
            .find(|cell| {
                cell.name == "Value"
                    && cell.locator.section.as_deref() == Some("Property")
                    && cell.locator.row.as_ref() == Some(row)
            })
            .and_then(|cell| cell.formula.clone())
    }

    #[test]
    fn shape_data_batch_is_one_undo_entry() {
        let session = session();
        let first = property_row(&session, "Device", None, "\"Amp\"", "Amp");
        let second = property_row(&session, "Owner", None, "\"Ada\"", "Ada");
        let before = session.undo_depth();
        let receipts = session
            .set_shape_data(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:1",
                &[
                    data_write(&first, "\"Mixer\""),
                    data_write(&second, "\"Grace\""),
                ],
            )
            .unwrap();
        assert!(
            receipts.iter().all(|receipt| !receipt.refused()),
            "{receipts:?}"
        );
        assert_eq!(
            value_formula(&session, &first).as_deref(),
            Some("\"Mixer\"")
        );
        assert_eq!(
            value_formula(&session, &second).as_deref(),
            Some("\"Grace\"")
        );
        assert_eq!(session.undo_depth(), before + 1);
        assert!(session.undo());
        assert_eq!(value_formula(&session, &first).as_deref(), Some("\"Amp\""));
        assert_eq!(value_formula(&session, &second).as_deref(), Some("\"Ada\""));
    }

    #[test]
    fn one_refused_row_leaves_the_whole_batch_unwritten() {
        let session = session();
        let writable = property_row(&session, "Device", None, "\"Amp\"", "Amp");
        let guarded = property_row(&session, "Serial", None, "GUARD(\"ABC\")", "ABC");
        let before = session.undo_depth();
        let receipts = session
            .set_shape_data(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:1",
                &[
                    data_write(&writable, "\"Mixer\""),
                    data_write(&guarded, "\"XYZ\""),
                ],
            )
            .unwrap();
        assert_eq!(receipts.len(), 2);
        assert!(!receipts[0].refused());
        assert!(receipts[0].after.is_none(), "{:?}", receipts[0]);
        let reason = receipts[1].refusal.as_deref().unwrap_or_default();
        assert!(reason.contains("GUARD"), "{reason}");
        assert_eq!(
            value_formula(&session, &writable).as_deref(),
            Some("\"Amp\"")
        );
        assert_eq!(
            value_formula(&session, &guarded).as_deref(),
            Some("GUARD(\"ABC\")")
        );
        assert_eq!(session.undo_depth(), before);
    }

    #[test]
    fn typed_rows_refuse_rather_than_losing_their_value() {
        let session = session();
        for (name, code, formula, value) in [
            ("Due", "5", "DATETIME(45000)", "3 March"),
            ("Runtime", "6", "DURATION(2)", "2 h"),
            ("Price", "7", "CY(4)", "4.00"),
        ] {
            let row = property_row(&session, name, Some(code), formula, value);
            let receipts = session
                .set_shape_data(
                    &EditCtx::local("a"),
                    "page:1",
                    "page:1:shape:1",
                    &[data_write(&row, "\"3 March\"")],
                )
                .unwrap();
            assert!(receipts[0].refused(), "{name}: {receipts:?}");
            assert_eq!(value_formula(&session, &row).as_deref(), Some(formula));
        }
    }

    #[test]
    fn a_number_row_refuses_a_non_numeric_value() {
        let session = session();
        let row = property_row(&session, "Count", Some("2"), "4", "4");
        let receipts = session
            .set_shape_data(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:1",
                &[data_write(&row, "\"seven\"")],
            )
            .unwrap();
        assert!(receipts[0].refused(), "{receipts:?}");
        assert_eq!(value_formula(&session, &row).as_deref(), Some("4"));
        let ok = session
            .set_shape_data(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:1",
                &[data_write(&row, "7")],
            )
            .unwrap();
        assert!(!ok[0].refused(), "{ok:?}");
        assert_eq!(value_formula(&session, &row).as_deref(), Some("7"));
    }

    #[test]
    fn an_empty_formula_is_rejected() {
        let session = session();
        let row = property_row(&session, "Device", None, "\"Amp\"", "Amp");
        let receipts = session
            .set_shape_data(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:1",
                &[data_write(&row, "   ")],
            )
            .unwrap();
        assert!(receipts[0].refused(), "{receipts:?}");
        assert_eq!(value_formula(&session, &row).as_deref(), Some("\"Amp\""));
    }

    #[test]
    fn guarded_property_value_refuses_edits() {
        let session = session();
        let row = CellRow::Name("Serial".to_owned());
        add_cell_at(
            &session,
            "Value",
            Some("Property"),
            Some(row.clone()),
            Some("GUARD(\"ABC\")"),
            Some("ABC"),
        );
        let locator = CellLocator {
            sheet: CellSheet::Page(1),
            shape_id: Some(1),
            section: Some("Property".to_owned()),
            section_index: None,
            row: Some(row),
            cell_name: "Value".to_owned(),
        };
        let error = session
            .set_cell_formula_at(
                &EditCtx::local("a"),
                "page:1",
                "page:1:shape:1",
                locator.clone(),
                "\"XYZ\"",
            )
            .expect_err("guarded shape-data values must refuse edits");
        assert!(error.to_string().contains("GUARD"));
        let snapshot = session.snapshot().unwrap();
        assert_eq!(
            snapshot.pages[0].shapes[0]
                .cells
                .iter()
                .find(|cell| cell.locator == locator)
                .unwrap()
                .formula
                .as_deref(),
            Some("GUARD(\"ABC\")")
        );
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
            master: None,
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
            master: None,
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
                    master: None,
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
            master: None,
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
            master: None,
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
    fn snapshot_carries_master_inherited_layer_member() {
        use vsdx_parse::{
            Cell, Row, RowChild, Section, SectionChild, Shape, ShapeChild, ShapesChild, Sheet,
            SheetChild,
        };
        fn cell(name: &str, value: &str) -> Cell {
            Cell {
                name: name.into(),
                formula: None,
                value: Some(value.into()),
                unit: None,
                del: false,
                other_attrs: Vec::new(),
            }
        }
        let mut package = vsdx_parse::parse_vsdx(include_bytes!(
            "../../vsdx-parse/tests/fixtures/foundation.vsdx"
        ))
        .unwrap();
        let page = package.page_part_paths[0].clone();
        let page_id = package.page_part_ids[&page];
        let contents = package.page_contents.get_mut(&page).unwrap();
        let mut shape_id = 0;
        for child in &mut contents.children {
            if let SheetChild::Shapes(shapes) = child
                && let Some(ShapesChild::Shape(shape)) = shapes.first_mut()
            {
                assert!(shape.cells().all(|cell| cell.name != "LayerMember"));
                shape.master = Some(999);
                shape_id = shape.id;
                break;
            }
        }
        assert_ne!(shape_id, 0);
        package
            .page_sheets
            .entry(page_id)
            .or_insert_with(|| Sheet {
                id: Some(page_id),
                children: Vec::new(),
                other_attrs: Vec::new(),
            })
            .children
            .push(SheetChild::Section(Section {
                name: "Layer".into(),
                index: None,
                del: false,
                children: vec![SectionChild::Row(Row {
                    index: Some(1),
                    name: None,
                    local_name: None,
                    row_type: None,
                    del: false,
                    children: vec![
                        RowChild::Cell(cell("Name", "Lighting")),
                        RowChild::Cell(cell("Visible", "0")),
                    ],
                    other_attrs: Vec::new(),
                })],
                other_attrs: Vec::new(),
            }));
        package
            .master_part_ids
            .insert("visio/masters/master999.xml".into(), 999);
        package.master_contents.insert(
            "visio/masters/master999.xml".into(),
            Sheet {
                id: None,
                children: vec![SheetChild::Shapes(vec![ShapesChild::Shape(Shape {
                    id: 1,
                    name: None,
                    name_u: None,
                    shape_type: None,
                    master: None,
                    master_shape: None,
                    line_style: None,
                    fill_style: None,
                    text_style: None,
                    children: vec![ShapeChild::Cell(cell("LayerMember", "1"))],
                    del: false,
                    other_attrs: Vec::new(),
                })])],
                other_attrs: Vec::new(),
            },
        );
        let resolved = vsdx_resolve::Resolver::new(&package)
            .resolve_shape(&page, shape_id)
            .unwrap();
        assert!(vsdx_resolve::shape_hidden_by_layers(
            &resolved,
            &vsdx_resolve::page_layers(&package, &page),
        ));
        let session = DiagramSession::from_package(package, 7).unwrap();
        let snapshot = session.snapshot().unwrap();
        let shape = snapshot.pages[0]
            .shapes
            .iter()
            .find(|shape| shape.source_id == shape_id)
            .unwrap();
        let member = shape
            .cells
            .iter()
            .find(|candidate| candidate.name == "LayerMember")
            .expect("seeded snapshot carries the master-inherited member");
        assert_eq!(member.value.as_deref(), Some("1"));
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
                            master: None,
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
                    master: None,
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

    fn control_plain(name: &str, formula: &str) -> CellSnapshot {
        CellSnapshot {
            locator: CellLocator {
                sheet: CellSheet::Page(0),
                shape_id: None,
                section: None,
                section_index: None,
                row: None,
                cell_name: name.to_owned(),
            },
            row_type: None,
            name: name.to_owned(),
            formula: Some(formula.to_owned()),
            value: None,
        }
    }

    fn control_handle_cell(name: &str, formula: &str) -> CellSnapshot {
        CellSnapshot {
            locator: CellLocator {
                sheet: CellSheet::Page(0),
                shape_id: None,
                section: Some("Control".to_owned()),
                section_index: None,
                row: Some(CellRow::Name("Row_1".to_owned())),
                cell_name: name.to_owned(),
            },
            row_type: None,
            name: name.to_owned(),
            formula: Some(formula.to_owned()),
            value: None,
        }
    }

    fn control_geometry_cell(
        index: u32,
        row_type: &str,
        name: &str,
        formula: &str,
    ) -> CellSnapshot {
        CellSnapshot {
            locator: CellLocator {
                sheet: CellSheet::Page(0),
                shape_id: None,
                section: Some("Geometry".to_owned()),
                section_index: None,
                row: Some(CellRow::Index(index)),
                cell_name: name.to_owned(),
            },
            row_type: Some(row_type.to_owned()),
            name: name.to_owned(),
            formula: Some(formula.to_owned()),
            value: None,
        }
    }

    fn control_draft() -> ShapeDraft {
        ShapeDraft {
            name: Some("Adjustable".to_owned()),
            master: None,
            cells: vec![
                control_plain("Width", "2"),
                control_plain("Height", "1"),
                control_plain("PinX", "3"),
                control_plain("PinY", "3"),
                control_plain("FillPattern", "1"),
                control_plain("FillForegnd", "RGB(1,2,3)"),
                control_plain("LinePattern", "1"),
                control_plain("LineColor", "RGB(4,5,6)"),
                control_plain("LineWeight", "0.02"),
                control_handle_cell("X", "Width*0.25"),
                control_handle_cell("Y", "Height*0.5"),
                control_handle_cell("XCon", "0"),
                control_handle_cell("YCon", "0"),
                control_geometry_cell(0, "MoveTo", "X", "0"),
                control_geometry_cell(0, "MoveTo", "Y", "0"),
                control_geometry_cell(1, "LineTo", "X", "Controls.Row_1.X"),
                control_geometry_cell(1, "LineTo", "Y", "0"),
                control_geometry_cell(2, "LineTo", "X", "Width-Controls.Row_1.X"),
                control_geometry_cell(2, "LineTo", "Y", "1"),
                control_geometry_cell(3, "LineTo", "X", "0"),
                control_geometry_cell(3, "LineTo", "Y", "1"),
            ],
        }
    }

    fn control_session() -> (DiagramSession, String, String) {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx");
        let session = DiagramSession::open(source, 1).unwrap();
        let page_id = session.snapshot().unwrap().pages[0].id.clone();
        let receipt = session
            .add_shape(&EditCtx::local("test"), &page_id, &control_draft())
            .unwrap();
        (session, page_id, receipt.shape_id)
    }

    fn control_locator(cell: &str) -> CellLocator {
        CellLocator {
            sheet: CellSheet::Page(0),
            shape_id: None,
            section: Some("Control".to_owned()),
            section_index: None,
            row: Some(CellRow::Name("Row_1".to_owned())),
            cell_name: cell.to_owned(),
        }
    }

    fn control_outline(session: &DiagramSession, part: &str, source_id: u32) -> Vec<(f64, f64)> {
        let package = session.package().unwrap();
        let list = vsdx_render::Renderer::default()
            .layout_page(&package, part)
            .unwrap();
        let expected = format!("{part}:{source_id}");
        let path = list
            .primitives
            .iter()
            .find_map(|primitive| match primitive {
                vsdx_render::Primitive::Shape { id, path, .. } if *id == expected => Some(path),
                _ => None,
            });
        let encoded = serde_json::to_string(&path.expect("control shape renders a path")).unwrap();
        serde_json::from_str::<Vec<serde_json::Value>>(&encoded)
            .unwrap()
            .into_iter()
            .map(|point| (point["x"].as_f64().unwrap(), point["y"].as_f64().unwrap()))
            .collect()
    }

    #[test]
    fn moving_a_control_handle_reshapes_the_rendered_outline() {
        let (session, page_id, shape_id) = control_session();
        let snapshot = session.snapshot().unwrap();
        let shape = snapshot.pages[0]
            .shapes
            .iter()
            .find(|shape| shape.id == shape_id)
            .unwrap();
        for (name, expected) in [("X", "0.5"), ("Y", "0.5")] {
            let cell = shape
                .cells
                .iter()
                .find(|cell| {
                    cell.locator.section.as_deref() == Some("Control")
                        && cell.locator.row == Some(CellRow::Name("Row_1".to_owned()))
                        && cell.name == name
                })
                .unwrap();
            assert_eq!(cell.value.as_deref(), Some(expected));
        }
        let part = snapshot.pages[0].source_part_path.clone();
        let source_id = shape.source_id;
        let before = control_outline(&session, &part, source_id);
        assert_eq!([(2.5, 2.5), (3.5, 3.5)], [before[1], before[2]]);
        session
            .set_cell_formula_at(
                &EditCtx::local("test"),
                &page_id,
                &shape_id,
                control_locator("X"),
                "1.5",
            )
            .unwrap();
        let after = control_outline(&session, &part, source_id);
        assert_eq!([(3.5, 2.5), (2.5, 3.5)], [after[1], after[2]]);
        let saved = session.save().unwrap();
        let reopened = DiagramSession::open(&saved, 2).unwrap();
        assert_eq!(control_outline(&reopened, &part, source_id), after);
    }

    #[test]
    fn guarded_control_cells_refuse_handle_edits() {
        let (session, page_id, shape_id) = control_session();
        session
            .set_cell_formula_at(
                &EditCtx::local("test"),
                &page_id,
                &shape_id,
                control_locator("X"),
                "GUARD(Width*0.25)",
            )
            .unwrap();
        let refused = session
            .set_cell_formula_at(
                &EditCtx::local("test"),
                &page_id,
                &shape_id,
                control_locator("X"),
                "1.5",
            )
            .unwrap_err();
        assert!(refused.to_string().contains("GUARD"));
    }

    fn control_formula(session: &DiagramSession, shape_id: &str, cell: &str) -> Option<String> {
        let snapshot = session.snapshot().unwrap();
        let shape = snapshot.pages[0]
            .shapes
            .iter()
            .find(|shape| shape.id == shape_id)?;
        shape
            .cells
            .iter()
            .find(|entry| {
                entry.locator.section.as_deref() == Some("Control")
                    && entry.locator.row == Some(CellRow::Name("Row_1".to_owned()))
                    && entry.name == cell
            })?
            .formula
            .clone()
    }

    #[test]
    fn a_control_handle_drag_writes_both_axes_as_one_undo_entry() {
        let (session, page_id, shape_id) = control_session();
        session.add_undo_barrier();
        let receipts = session
            .set_control_handle(
                &EditCtx::local("test"),
                &page_id,
                &shape_id,
                "Row_1",
                Some("1.5".to_owned()),
                Some("0.25".to_owned()),
            )
            .unwrap();
        assert_eq!(receipts.len(), 2);
        assert_eq!(
            [
                control_formula(&session, &shape_id, "X"),
                control_formula(&session, &shape_id, "Y")
            ],
            [Some("1.5".to_owned()), Some("0.25".to_owned())]
        );
        assert!(session.undo());
        assert_eq!(
            [
                control_formula(&session, &shape_id, "X"),
                control_formula(&session, &shape_id, "Y")
            ],
            [Some("Width*0.25".to_owned()), Some("Height*0.5".to_owned())]
        );
    }

    #[test]
    fn a_guarded_axis_refuses_the_whole_control_handle_drag() {
        let (session, page_id, shape_id) = control_session();
        session
            .set_cell_formula_at(
                &EditCtx::local("test"),
                &page_id,
                &shape_id,
                control_locator("X"),
                "GUARD(Width*0.25)",
            )
            .unwrap();
        let refused = session
            .set_control_handle(
                &EditCtx::local("test"),
                &page_id,
                &shape_id,
                "Row_1",
                Some("1.5".to_owned()),
                Some("0.25".to_owned()),
            )
            .unwrap_err();
        assert!(refused.to_string().contains("GUARD"));
        assert_eq!(
            control_formula(&session, &shape_id, "Y"),
            Some("Height*0.5".to_owned())
        );
        let receipts = session
            .set_control_handle(
                &EditCtx::local("test"),
                &page_id,
                &shape_id,
                "Row_1",
                None,
                Some("0.25".to_owned()),
            )
            .unwrap();
        assert_eq!(receipts.len(), 1);
        assert_eq!(
            [
                control_formula(&session, &shape_id, "X"),
                control_formula(&session, &shape_id, "Y")
            ],
            [
                Some("GUARD(Width*0.25)".to_owned()),
                Some("0.25".to_owned())
            ]
        );
    }

    fn route_fixture() -> DiagramSession {
        DiagramSession::open(
            include_bytes!("../../vsdx-parse/tests/fixtures/connector-route-style.vsdx"),
            31,
        )
        .unwrap()
    }

    fn connector_draft() -> ShapeDraft {
        ShapeDraft {
            name: Some("Connector".to_owned()),
            master: None,
            cells: [
                ("OneD", "1"),
                ("BeginX", "1"),
                ("BeginY", "2"),
                ("EndX", "4"),
                ("EndY", "2"),
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
            .collect(),
        }
    }

    fn grouped_glue_session() -> DiagramSession {
        DiagramSession::open(
            include_bytes!("../../vsdx-parse/tests/fixtures/grouped-glue.vsdx"),
            901,
        )
        .unwrap()
    }

    fn route_cells(session: &DiagramSession, shape: &str) -> Vec<(u32, String, Option<String>)> {
        let mut cells = session.snapshot().unwrap().pages[0]
            .shapes
            .iter()
            .find(|candidate| candidate.id == shape)
            .unwrap()
            .cells
            .iter()
            .filter(|cell| cell.locator.section.as_deref() == Some("Geometry"))
            .map(|cell| {
                let row = match cell.locator.row {
                    Some(CellRow::Index(index)) => index,
                    _ => u32::MAX,
                };
                (row, cell.locator.cell_name.clone(), cell.formula.clone())
            })
            .collect::<Vec<_>>();
        cells.sort();
        cells
    }

    #[test]
    fn set_connector_route_rewrites_filed_geometry() {
        let session = route_fixture();
        let receipt = session
            .set_connector_route(
                &EditCtx::local("test"),
                "page:1",
                "page:1:shape:1",
                &[(1.0, 1.0), (2.5, 1.0), (2.5, 3.0), (4.0, 3.0)],
            )
            .unwrap();
        assert_eq!(receipt.points, 4);
        assert_eq!(
            route_cells(&session, "page:1:shape:1"),
            vec![
                (0, "X".to_owned(), Some("0".to_owned())),
                (0, "Y".to_owned(), Some("0".to_owned())),
                (1, "X".to_owned(), Some("1.5".to_owned())),
                (1, "Y".to_owned(), Some("0".to_owned())),
                (2, "X".to_owned(), Some("1.5".to_owned())),
                (2, "Y".to_owned(), Some("2".to_owned())),
                (3, "X".to_owned(), Some("3".to_owned())),
                (3, "Y".to_owned(), Some("2".to_owned())),
            ]
        );
        let saved = session.save().unwrap();
        let reparsed = vsdx_parse::parse_vsdx(&saved).unwrap();
        let shape = reparsed.page_contents["visio/pages/page1.xml"]
            .shapes()
            .next()
            .unwrap();
        let rows = shape
            .sections()
            .find(|section| section.name == "Geometry")
            .unwrap()
            .rows()
            .map(|row| {
                (
                    row.row_type.clone(),
                    row.cells()
                        .map(|cell| (cell.name.clone(), cell.formula.clone()))
                        .collect::<Vec<_>>(),
                )
            })
            .collect::<Vec<_>>();
        assert_eq!(rows.len(), 4);
        assert_eq!(rows[0].0.as_deref(), Some("MoveTo"));
        assert!(
            rows[1..]
                .iter()
                .all(|row| row.0.as_deref() == Some("LineTo"))
        );
    }

    fn nested_groups_session() -> DiagramSession {
        DiagramSession::open(
            include_bytes!("../../vsdx-parse/tests/fixtures/nested-groups.vsdx"),
            902,
        )
        .unwrap()
    }

    fn find_by_source(shapes: &[ShapeSnapshot], source_id: u32) -> &ShapeSnapshot {
        shapes
            .iter()
            .find_map(|shape| {
                if shape.source_id == source_id {
                    Some(shape)
                } else {
                    find_by_source_result(&shape.children, source_id)
                }
            })
            .expect("source shape is part of the snapshot")
    }

    fn find_by_source_result(shapes: &[ShapeSnapshot], source_id: u32) -> Option<&ShapeSnapshot> {
        shapes.iter().find_map(|shape| {
            if shape.source_id == source_id {
                Some(shape)
            } else {
                find_by_source_result(&shape.children, source_id)
            }
        })
    }

    fn test_page_number(page_id: &str) -> u32 {
        page_id
            .strip_prefix("page:")
            .and_then(|value| value.parse::<u32>().ok())
            .expect("test page ID carries its source page")
    }

    fn tree_draft(
        session: &DiagramSession,
        page_id: &str,
        shape: &ShapeSnapshot,
    ) -> ShapeTreeDraft {
        let mut draft = tree_node(session, page_id, shape);
        draft.glue = session.subtree_glue(page_id, &shape.id).unwrap();
        draft
    }

    fn tree_node(session: &DiagramSession, page_id: &str, shape: &ShapeSnapshot) -> ShapeTreeDraft {
        ShapeTreeDraft {
            name: shape.name.clone(),
            cells: shape.cells.clone(),
            text: session.shape_text(page_id, &shape.id).unwrap(),
            copy_source_id: Some(shape.copy_source_id.unwrap_or(shape.source_id)),
            copy_source_page_id: Some(
                shape
                    .copy_source_page_id
                    .unwrap_or(test_page_number(page_id)),
            ),
            source_shape_id: Some(shape.id.clone()),
            source_id: Some(shape.source_id),
            copy_refusal: shape.copy_refusal.clone(),
            glue: Vec::new(),
            children: shape
                .children
                .iter()
                .map(|child| tree_node(session, page_id, child))
                .collect(),
        }
    }

    fn paste_group(
        session: &DiagramSession,
        page_id: &str,
        source_id: u32,
    ) -> (ShapeSnapshot, ShapeReceipt) {
        let snapshot = session.snapshot().unwrap();
        let source = find_by_source(&snapshot.pages[0].shapes, source_id).clone();
        let draft = tree_draft(session, page_id, &source);
        let receipt = session
            .add_shape_tree(&EditCtx::local("paste"), page_id, &draft)
            .unwrap();
        let snapshot = session.snapshot().unwrap();
        let pasted = snapshot.pages[0]
            .shapes
            .iter()
            .find(|shape| shape.id == receipt.shape_id)
            .cloned()
            .unwrap();
        (pasted, receipt)
    }

    fn nest_shape(session: &DiagramSession, parent_id: &str, child_id: &str) {
        let mut txn = session.yrs_doc().transact_mut();
        let pages = txn.get_map(PAGES).unwrap();
        let page = match pages.get(&txn, "page:1").unwrap() {
            yrs::Out::YMap(page) => page,
            _ => unreachable!(),
        };
        let roots = match page.get(&txn, "shapes").unwrap() {
            yrs::Out::YArray(roots) => roots,
            _ => unreachable!(),
        };
        let from = (0..roots.len(&txn))
            .find(|index| match roots.get(&txn, *index) {
                Some(yrs::Out::Any(Any::String(id))) => id.to_string() == child_id,
                _ => false,
            })
            .unwrap();
        roots.remove_range(&mut txn, from, 1);
        let sheets = txn.get_map(SHEETS).unwrap();
        let child = match sheets.get(&txn, child_id).unwrap() {
            yrs::Out::YMap(child) => child,
            _ => unreachable!(),
        };
        child.insert(&mut txn, "parentId", parent_id);
        let parent = match sheets.get(&txn, parent_id).unwrap() {
            yrs::Out::YMap(parent) => parent,
            _ => unreachable!(),
        };
        let siblings = match parent.get(&txn, "shapes").unwrap() {
            yrs::Out::YArray(siblings) => siblings,
            _ => unreachable!(),
        };
        siblings.push_back(&mut txn, child_id);
    }

    fn cell_formula<'a>(shape: &'a ShapeSnapshot, name: &str) -> Option<&'a str> {
        shape
            .cells
            .iter()
            .find(|cell| cell.name == name)
            .and_then(|cell| cell.formula.as_deref())
    }

    fn collected_formulas(shape: &ShapeSnapshot) -> Vec<String> {
        let mut formulas = shape
            .cells
            .iter()
            .filter_map(|cell| cell.formula.clone())
            .collect::<Vec<_>>();
        for child in &shape.children {
            formulas.extend(collected_formulas(child));
        }
        formulas
    }

    #[test]
    fn group_paste_reproduces_children_at_relative_positions() {
        let session = grouped_glue_session();
        let before = session.snapshot().unwrap().pages[0].shapes.len();
        let source = find_by_source(&session.snapshot().unwrap().pages[0].shapes, 10).clone();
        let (pasted, _) = paste_group(&session, "page:1", 10);
        assert_eq!(
            session.snapshot().unwrap().pages[0].shapes.len(),
            before + 1
        );
        assert_ne!(pasted.source_id, 10);
        assert_eq!(pasted.children.len(), 1);
        assert!(pasted.children[0].children.is_empty());
        for (original, copy) in [
            (&source, &pasted),
            (&source.children[0], &pasted.children[0]),
        ] {
            assert_eq!(original.cells.len(), copy.cells.len());
            for cell in &original.cells {
                let mirror = copy
                    .cells
                    .iter()
                    .find(|candidate| {
                        candidate.name == cell.name
                            && candidate.locator.section == cell.locator.section
                            && candidate.locator.section_index == cell.locator.section_index
                            && candidate.locator.row == cell.locator.row
                    })
                    .unwrap();
                assert_eq!(mirror.formula, cell.formula);
                assert_eq!(mirror.value, cell.value);
            }
        }
        let reopened = DiagramSession::open(&session.save().unwrap(), 905).unwrap();
        assert_reopened_projection_eq(&session, &reopened, "group paste");
    }

    #[test]
    fn pasted_tree_references_itself_never_the_original() {
        let session = grouped_glue_session();
        let snapshot = session.snapshot().unwrap();
        let child_id = find_by_source(&snapshot.pages[0].shapes, 11).id.clone();
        session
            .set_cell_formula(
                &EditCtx::local("edit"),
                "page:1",
                &child_id,
                "PinX",
                "Sheet.10!Width/2",
            )
            .unwrap();
        let (pasted, _) = paste_group(&session, "page:1", 10);
        let expected = format!("Sheet.{}!Width/2", pasted.source_id);
        assert_eq!(
            cell_formula(&pasted.children[0], "PinX"),
            Some(expected.as_str())
        );
        let saved = session.save().unwrap();
        assert_eq!(session.save().unwrap(), saved);
        let reopened = DiagramSession::open(&saved, 906).unwrap();
        assert_reopened_projection_eq(&session, &reopened, "self-referential paste");
        let resnapshot = reopened.snapshot().unwrap();
        let root = find_by_source(&resnapshot.pages[0].shapes, pasted.source_id);
        assert_eq!(
            cell_formula(&root.children[0], "PinX"),
            Some(expected.as_str())
        );
        let original = find_by_source(&resnapshot.pages[0].shapes, 10);
        assert_eq!(
            cell_formula(&original.children[0], "PinX"),
            Some("Sheet.10!Width/2")
        );
        assert!(
            collected_formulas(root)
                .iter()
                .all(|formula| !formula.contains("Sheet.10!")),
            "pasted tree points back at the original"
        );
    }

    #[test]
    fn set_connector_route_round_trips_through_peers() {
        let left = route_fixture();
        let right =
            DiagramSession::open_from_update(&left.encode_state_as_update_v1(), 32).unwrap();
        left.set_connector_route(
            &EditCtx::local("left"),
            "page:1",
            "page:1:shape:1",
            &[(1.0, 1.0), (4.0, 1.0), (4.0, 3.0)],
        )
        .unwrap();
        let update = left
            .encode_diff_v1(&right.encode_state_vector_v1())
            .unwrap();
        right.apply_update_v1(&update).unwrap();
        assert_eq!(
            route_cells(&left, "page:1:shape:1"),
            route_cells(&right, "page:1:shape:1")
        );
    }

    #[test]
    fn set_connector_route_collapses_surplus_rows() {
        let session = route_fixture();
        let receipt = session
            .set_connector_route(
                &EditCtx::local("test"),
                "page:1",
                "page:1:shape:1",
                &[(1.0, 1.0), (2.0, 2.0), (3.0, 2.0), (4.0, 3.0), (4.0, 3.0)],
            )
            .unwrap();
        assert_eq!(receipt.points, 4);
        session
            .set_connector_route(
                &EditCtx::local("test"),
                "page:1",
                "page:1:shape:1",
                &[(1.0, 1.0), (4.0, 3.0)],
            )
            .unwrap();
        let cells = route_cells(&session, "page:1:shape:1");
        assert_eq!(cells.len(), 8);
        for (_, _, formula) in cells.iter().skip(2) {
            assert!(formula.as_deref() == Some("3") || formula.as_deref() == Some("2"));
        }
    }

    /// The renderer appends the glued endpoint, so a re-fed route must not grow a row per drag.
    #[test]
    fn set_connector_route_drops_a_repeated_endpoint() {
        let session = route_fixture();
        session
            .set_connector_route(
                &EditCtx::local("test"),
                "page:1",
                "page:1:shape:1",
                &[(1.0, 1.0), (2.5, 1.0), (4.0, 3.0), (4.0, 3.0)],
            )
            .unwrap();
        assert_eq!(
            route_cells(&session, "page:1:shape:1"),
            vec![
                (0, "X".to_owned(), Some("0".to_owned())),
                (0, "Y".to_owned(), Some("0".to_owned())),
                (1, "X".to_owned(), Some("1.5".to_owned())),
                (1, "Y".to_owned(), Some("0".to_owned())),
                (2, "X".to_owned(), Some("3".to_owned())),
                (2, "Y".to_owned(), Some("2".to_owned())),
            ]
        );
        assert_eq!(
            session
                .set_connector_route(
                    &EditCtx::local("test"),
                    "page:1",
                    "page:1:shape:1",
                    &[(2.0, 2.0), (2.0, 2.0)],
                )
                .unwrap_err()
                .to_string(),
            "invalid diagram state: connector route needs two distinct points"
        );
    }

    /// Page 4 of the fixture files two subpaths; a hand route must refuse, not overwrite them.
    #[test]
    fn set_connector_route_refuses_a_multi_subpath_geometry() {
        let session = route_fixture();
        assert_eq!(
            session
                .set_connector_route(
                    &EditCtx::local("test"),
                    "page:4",
                    "page:4:shape:1",
                    &[(1.0, 1.0), (4.0, 3.0)],
                )
                .unwrap_err()
                .to_string(),
            "invalid diagram state: connector Geometry is not a single straight run"
        );
    }

    /// A refusal in a later row must leave every earlier row untouched.
    #[test]
    fn set_connector_route_refuses_before_the_first_write() {
        let guarded = session();
        for name in ["OneD", "BeginX", "BeginY", "EndX", "EndY"] {
            add_cell(&guarded, name, Some("1"), None);
        }
        for name in ["PinX", "PinY", "Width", "Height"] {
            add_cell(&guarded, name, Some("1"), None);
        }
        for (row, name, formula) in [
            (0, "X", "9"),
            (0, "Y", "9"),
            (1, "X", "GUARD(0)"),
            (1, "Y", "9"),
        ] {
            add_cell_at(
                &guarded,
                name,
                Some("Geometry"),
                Some(CellRow::Index(row)),
                Some(formula),
                None,
            );
        }
        assert_eq!(
            guarded
                .set_connector_route(
                    &EditCtx::local("test"),
                    "page:1",
                    "page:1:shape:1",
                    &[(0.0, 0.0), (1.0, 1.0)],
                )
                .unwrap_err()
                .to_string(),
            "invalid diagram state: GUARD protects the requested cell"
        );
        assert_eq!(
            route_cells(&guarded, "page:1:shape:1")
                .into_iter()
                .map(|(_, _, formula)| formula)
                .collect::<Vec<_>>(),
            vec![
                Some("9".to_owned()),
                Some("9".to_owned()),
                Some("GUARD(0)".to_owned()),
                Some("9".to_owned()),
            ]
        );
    }

    #[test]
    fn glue_inside_the_subtree_survives_remapped() {
        let session = grouped_glue_session();
        let (first, _) = paste_group(&session, "page:1", 10);
        let child_id = first.children[0].id.clone();
        let connector = session
            .add_connector(
                &EditCtx::local("connector"),
                "page:1",
                &connector_draft(),
                &ConnectorGlue {
                    shape_id: first.id.clone(),
                    to_cell: None,
                },
                &ConnectorGlue {
                    shape_id: child_id,
                    to_cell: Some("Connections.X1".to_owned()),
                },
            )
            .unwrap();
        nest_shape(&session, &first.id, &connector.shape_id);
        let snapshot = session.snapshot().unwrap();
        let nested = snapshot.pages[0]
            .shapes
            .iter()
            .find(|shape| shape.id == first.id)
            .cloned()
            .unwrap();
        assert_eq!(nested.children.len(), 2);
        let draft = tree_draft(&session, "page:1", &nested);
        let receipt = session
            .add_shape_tree(&EditCtx::local("paste"), "page:1", &draft)
            .unwrap();
        let snapshot = session.snapshot().unwrap();
        let pasted = snapshot.pages[0]
            .shapes
            .iter()
            .find(|shape| shape.id == receipt.shape_id)
            .unwrap();
        assert_eq!(pasted.children.len(), 2);
        let pasted_connector = pasted
            .children
            .iter()
            .find(|child| child.name.as_deref() == Some("Connector"))
            .unwrap();
        let package = session.package().unwrap();
        let part = package.page_part_paths[0].clone();
        let glue = package.page_contents[&part]
            .connects()
            .filter(|connect| connect.from_sheet == pasted_connector.source_id)
            .collect::<Vec<_>>();
        assert_eq!(glue.len(), 2);
        let targets = glue
            .iter()
            .map(|connect| connect.to_sheet)
            .collect::<std::collections::HashSet<_>>();
        assert_eq!(targets.len(), 2);
        assert!(targets.contains(&pasted.source_id));
        assert!(targets.contains(&pasted.children[0].source_id));
        let saved = session.save().unwrap();
        let reopened = DiagramSession::open(&saved, 907).unwrap();
        let live_package = session.package().unwrap();
        let reopened_package = reopened.package().unwrap();
        let renderer = vsdx_render::Renderer::default();
        assert_eq!(
            renderer.layout_page(&live_package, &part).unwrap(),
            renderer.layout_page(&reopened_package, &part).unwrap(),
            "pasted internal glue rendering"
        );
        let resnapshot = reopened.snapshot().unwrap();
        let rerooted = find_by_source(&resnapshot.pages[0].shapes, pasted.source_id).clone();
        assert_eq!(rerooted.children.len(), 2);
        let reglued = reopened_package.page_contents[&part]
            .connects()
            .filter(|connect| connect.from_sheet == pasted_connector.source_id)
            .collect::<Vec<_>>();
        assert_eq!(reglued.len(), 2);
        let connectivity = vsdx_resolve::Resolver::new(&reopened_package)
            .resolve_page_connectivity(&part)
            .unwrap();
        let resolved = &connectivity.connectors[&pasted_connector.source_id];
        assert_eq!(resolved.glue.len(), 2);
        assert!(resolved.glue.iter().all(|glue| {
            glue.to
                .as_ref()
                .is_some_and(|end| end.connection_point.is_some())
        }));
        assert!(connectivity.diagnostics.is_empty());
    }

    #[test]
    fn glue_crossing_the_copy_boundary_is_dropped() {
        let session = grouped_glue_session();
        let (pasted, _) = paste_group(&session, "page:1", 10);
        let package = session.package().unwrap();
        let part = package.page_part_paths[0].clone();
        let connects = package.page_contents[&part].connects().collect::<Vec<_>>();
        assert_eq!(connects.len(), 6);
        assert!(connects.iter().all(|connect| connect.from_sheet == 1));
        let snapshot = session.snapshot().unwrap();
        let connector = find_by_source(&snapshot.pages[0].shapes, 1).clone();
        let receipt = session
            .add_shape_tree(
                &EditCtx::local("paste"),
                "page:1",
                &tree_draft(&session, "page:1", &connector),
            )
            .unwrap();
        let resnapshot = session.snapshot().unwrap();
        let copy = resnapshot.pages[0]
            .shapes
            .iter()
            .find(|shape| shape.id == receipt.shape_id)
            .unwrap();
        let package = session.package().unwrap();
        assert_eq!(package.page_contents[&part].connects().count(), 6);
        let connectivity = vsdx_resolve::Resolver::new(&package)
            .resolve_page_connectivity(&part)
            .unwrap();
        assert!(connectivity.connectors[&copy.source_id].glue.is_empty());
        assert_eq!(connectivity.connectors[&1].glue.len(), 6);
        assert_eq!(pasted.children.len(), 1);
        let saved = session.save().unwrap();
        let reopened = DiagramSession::open(&saved, 908).unwrap();
        assert_reopened_projection_eq(&session, &reopened, "dropped boundary glue");
    }

    #[test]
    fn group_paste_is_one_undo_step() {
        let session = grouped_glue_session();
        let before = session.snapshot().unwrap().pages[0].shapes.len();
        paste_group(&session, "page:1", 30);
        assert_eq!(
            session.snapshot().unwrap().pages[0].shapes.len(),
            before + 1
        );
        assert!(session.undo());
        assert_eq!(session.snapshot().unwrap().pages[0].shapes.len(), before);
        assert!(session.redo());
        assert_eq!(
            session.snapshot().unwrap().pages[0].shapes.len(),
            before + 1
        );
    }

    #[test]
    fn group_paste_syncs_to_a_peer_as_one_update() {
        let session = grouped_glue_session();
        let before = session.encode_state_as_update_v1();
        paste_group(&session, "page:1", 10);
        let peer = DiagramSession::open_from_update(&before, 909).unwrap();
        let diff = session
            .encode_diff_v1(&peer.encode_state_vector_v1())
            .unwrap();
        peer.apply_update_v1(&diff).unwrap();
        assert_eq!(peer.snapshot().unwrap(), session.snapshot().unwrap());
        let saved = peer.save().unwrap();
        let reopened = DiagramSession::open(&saved, 910).unwrap();
        assert_reopened_projection_eq(&peer, &reopened, "synced group paste");
    }

    #[test]
    fn opening_and_saving_without_an_edit_keeps_every_part_byte_identical() {
        for bytes in [
            include_bytes!("../../vsdx-parse/tests/fixtures/grouped-glue.vsdx").as_slice(),
            include_bytes!("../../vsdx-parse/tests/fixtures/nested-groups.vsdx").as_slice(),
        ] {
            let original = vsdx_parse::parse_vsdx(bytes).unwrap();
            let saved = DiagramSession::open(bytes, 913).unwrap().save().unwrap();
            let reparsed = vsdx_parse::parse_vsdx(&saved).unwrap();
            let paths = [
                vec![original.document_part_path.clone()],
                original.pages_part_path.iter().cloned().collect(),
                original.masters_part_path.iter().cloned().collect(),
                original.windows_part_path.iter().cloned().collect(),
                original.page_part_paths.clone(),
                original.master_part_paths.clone(),
                original.theme_part_paths.clone(),
            ]
            .concat();
            for path in paths {
                assert_eq!(
                    reparsed.part_bytes(&path),
                    original.part_bytes(&path),
                    "part changed without an edit: {path}"
                );
            }
        }
    }

    #[test]
    fn a_group_paste_undone_saves_byte_identically() {
        let session = grouped_glue_session();
        let before = session.save().unwrap();
        paste_group(&session, "page:1", 30);
        assert!(session.undo());
        assert_eq!(session.save().unwrap(), before);
    }

    #[test]
    fn group_paste_survives_save_and_reopen_with_untouched_parts_intact() {
        let session = grouped_glue_session();
        let snapshot = session.snapshot().unwrap();
        let child_id = find_by_source(&snapshot.pages[0].shapes, 32).id.clone();
        session
            .set_shape_text(&EditCtx::local("text"), "page:1", &child_id, "nested hello")
            .unwrap();
        let (pasted, _) = paste_group(&session, "page:1", 30);
        assert_eq!(
            session
                .shape_text("page:1", &pasted.children[0].children[0].id)
                .unwrap(),
            "nested hello"
        );
        let saved = session.save().unwrap();
        assert_eq!(session.save().unwrap(), saved);
        let first = vsdx_parse::parse_vsdx(&saved).unwrap();
        let fixture = grouped_glue_session().package().unwrap();
        for path in &fixture.page_part_paths {
            if path.ends_with("page1.xml") {
                continue;
            }
            assert_eq!(
                first.part_bytes(path),
                fixture.part_bytes(path),
                "untouched part changed: {path}"
            );
        }
        let reopened = DiagramSession::open(&saved, 911).unwrap();
        assert_reopened_projection_eq(&session, &reopened, "saved group paste");
        let resnapshot = reopened.snapshot().unwrap();
        let root = find_by_source(&resnapshot.pages[0].shapes, pasted.source_id);
        assert_eq!(root.children.len(), 1);
        assert_eq!(root.children[0].children.len(), 1);
        assert_eq!(
            reopened
                .shape_text("page:1", &root.children[0].children[0].id)
                .unwrap(),
            "nested hello"
        );
    }

    #[test]
    fn set_connector_route_refuses_shapes_without_endpoints() {
        let session = DiagramSession::open(
            include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx"),
            33,
        )
        .unwrap();
        assert_eq!(
            session
                .set_connector_route(
                    &EditCtx::local("test"),
                    "page:1",
                    "page:1:shape:1",
                    &[(0.0, 0.0), (1.0, 1.0)],
                )
                .unwrap_err()
                .to_string(),
            "invalid diagram state: shape is not a 1D connector"
        );
        assert_eq!(
            session
                .set_connector_route(
                    &EditCtx::local("test"),
                    "page:1",
                    "page:1:shape:1",
                    &[(0.0, 0.0)],
                )
                .unwrap_err()
                .to_string(),
            "invalid diagram state: connector route needs between 2 and 256 points"
        );
    }

    #[test]
    fn set_connector_route_honours_guarded_geometry() {
        let guarded = session();
        for name in ["OneD", "BeginX", "BeginY", "EndX", "EndY"] {
            add_cell(&guarded, name, Some("1"), None);
        }
        for name in ["PinX", "PinY", "Width", "Height"] {
            add_cell(&guarded, name, Some("1"), None);
        }
        add_cell_at(
            &guarded,
            "X",
            Some("Geometry"),
            Some(CellRow::Index(0)),
            Some("GUARD(0)"),
            None,
        );
        assert_eq!(
            guarded
                .set_connector_route(
                    &EditCtx::local("test"),
                    "page:1",
                    "page:1:shape:1",
                    &[(0.0, 0.0), (1.0, 1.0)],
                )
                .unwrap_err()
                .to_string(),
            "invalid diagram state: GUARD protects the requested cell"
        );
    }

    #[test]
    fn group_copy_refuses_unportable_content_and_leaves_no_trace() {
        let session = nested_groups_session();
        let snapshot = session.snapshot().unwrap();
        let before = session.encode_state_as_update_v1();
        for source_id in [1, 2] {
            let source = find_by_source(&snapshot.pages[0].shapes, source_id).clone();
            let error = session
                .add_shape_tree(
                    &EditCtx::local("paste"),
                    "page:1",
                    &tree_draft(&session, "page:1", &source),
                )
                .unwrap_err();
            assert!(
                error.to_string().contains("embedded media"),
                "unexpected refusal for {source_id}: {error}"
            );
        }
        let leaf = find_by_source(&snapshot.pages[0].shapes, 4).clone();
        let error = session
            .add_shape_tree(
                &EditCtx::local("paste"),
                "page:1",
                &tree_draft(&session, "page:1", &leaf),
            )
            .unwrap_err();
        assert!(
            error.to_string().contains("embedded media"),
            "unexpected leaf refusal: {error}"
        );
        assert_eq!(session.encode_state_as_update_v1(), before);
        let clean = find_by_source(&snapshot.pages[0].shapes, 3).clone();
        session
            .add_shape_tree(
                &EditCtx::local("paste"),
                "page:1",
                &tree_draft(&session, "page:1", &clean),
            )
            .unwrap();
        let reopened = DiagramSession::open(&session.save().unwrap(), 912).unwrap();
        assert_reopened_projection_eq(&session, &reopened, "leaf paste beside refusal");
    }

    #[test]
    fn group_paste_refuses_malformed_drafts_atomically() {
        let session = grouped_glue_session();
        let before = session.encode_state_as_update_v1();
        let mut deep = ShapeTreeDraft {
            name: None,
            cells: Vec::new(),
            text: String::new(),
            copy_source_id: None,
            copy_source_page_id: None,
            source_shape_id: Some("page:1:shape:10".to_owned()),
            source_id: Some(10),
            copy_refusal: None,
            glue: Vec::new(),
            children: Vec::new(),
        };
        for _ in 0..crate::diagram::MAX_SHAPE_NESTING {
            deep = ShapeTreeDraft {
                name: None,
                cells: Vec::new(),
                text: String::new(),
                copy_source_id: None,
                copy_source_page_id: None,
                source_shape_id: Some("page:1:shape:10".to_owned()),
                source_id: Some(10),
                copy_refusal: None,
                glue: Vec::new(),
                children: vec![deep],
            };
        }
        let error = session
            .add_shape_tree(&EditCtx::local("paste"), "page:1", &deep)
            .unwrap_err();
        assert!(
            error.to_string().contains("maximum depth"),
            "unexpected depth refusal: {error}"
        );
        let snapshot = session.snapshot().unwrap();
        let source = find_by_source(&snapshot.pages[0].shapes, 10).clone();
        let mut duplicated = tree_draft(&session, "page:1", &source);
        duplicated.children.push(duplicated.children[0].clone());
        let error = session
            .add_shape_tree(&EditCtx::local("paste"), "page:1", &duplicated)
            .unwrap_err();
        assert!(
            error.to_string().contains("twice"),
            "unexpected duplicate refusal: {error}"
        );
        assert_eq!(session.encode_state_as_update_v1(), before);
    }

    #[test]
    fn paste_survives_a_source_child_deleted_after_the_copy() {
        let session = grouped_glue_session();
        let snapshot = session.snapshot().unwrap();
        let source = find_by_source(&snapshot.pages[0].shapes, 10).clone();
        let draft = tree_draft(&session, "page:1", &source);
        session
            .delete_shape(&EditCtx::local("delete"), "page:1", &source.children[0].id)
            .unwrap();
        let receipt = session
            .add_shape_tree(&EditCtx::local("paste"), "page:1", &draft)
            .unwrap();
        let snapshot = session.snapshot().unwrap();
        let pasted = snapshot.pages[0]
            .shapes
            .iter()
            .find(|shape| shape.id == receipt.shape_id)
            .unwrap();
        assert_eq!(pasted.children.len(), source.children.len());
    }

    #[test]
    fn a_cross_page_paste_refuses_a_reference_it_did_not_copy() {
        let leaked = |page_id: &str| {
            let session = session();
            let draft = ShapeTreeDraft {
                name: None,
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
                    formula: Some("Sheet.2!Width".to_owned()),
                    value: None,
                }],
                text: String::new(),
                copy_source_id: Some(1),
                copy_source_page_id: Some(1),
                source_shape_id: Some("page:1:shape:1".to_owned()),
                source_id: Some(1),
                copy_refusal: None,
                glue: Vec::new(),
                children: Vec::new(),
            };
            session.add_shape_tree(&EditCtx::local("paste"), page_id, &draft)
        };
        assert!(leaked("page:1").is_ok());
        let error = leaked("page:2").unwrap_err();
        assert!(
            error.to_string().contains("outside the copy"),
            "unexpected cross-page refusal: {error}"
        );
    }
    #[test]
    fn cutting_a_group_agrees_with_copy_on_boundary_glue() {
        let session = grouped_glue_session();
        let snapshot = session.snapshot().unwrap();
        let source = find_by_source(&snapshot.pages[0].shapes, 10).clone();
        let draft = tree_draft(&session, "page:1", &source);
        session
            .delete_shape(&EditCtx::local("cut"), "page:1", &source.id)
            .unwrap();
        session
            .add_shape_tree(&EditCtx::local("paste"), "page:1", &draft)
            .unwrap();
        let package = session.package().unwrap();
        let part = package.page_part_paths[0].clone();
        let connects = package.page_contents[&part].connects().collect::<Vec<_>>();
        assert_eq!(connects.len(), 4);
        assert!(
            connects
                .iter()
                .all(|connect| connect.from_sheet == 1 && [21, 32].contains(&connect.to_sheet))
        );
        let saved = session.save().unwrap();
        let reopened = DiagramSession::open(&saved, 913).unwrap();
        assert_reopened_projection_eq(&session, &reopened, "cut group paste");
    }

    #[test]
    fn seeded_package_glue_survives_group_copy() {
        let session = grouped_glue_session();
        let snapshot = session.snapshot().unwrap();
        let connector = find_by_source(&snapshot.pages[0].shapes, 1).id.clone();
        let group = find_by_source(&snapshot.pages[0].shapes, 10).id.clone();
        nest_shape(&session, &group, &connector);
        let target = find_by_source(&session.snapshot().unwrap().pages[0].shapes, 11)
            .id
            .clone();
        let glue = session.subtree_glue("page:1", &group).unwrap();
        assert_eq!(glue.len(), 1);
        assert_eq!(glue[0].connector_source, connector);
        assert_eq!(glue[0].endpoint, "begin");
        assert_eq!(glue[0].target_source, target);
        assert_eq!(glue[0].to_cell, "Connections.X1");
        let snapshot = session.snapshot().unwrap();
        let source = find_by_source(&snapshot.pages[0].shapes, 10).clone();
        let draft = tree_draft(&session, "page:1", &source);
        let receipt = session
            .add_shape_tree(&EditCtx::local("paste"), "page:1", &draft)
            .unwrap();
        let resnapshot = session.snapshot().unwrap();
        let pasted = resnapshot.pages[0]
            .shapes
            .iter()
            .find(|shape| shape.id == receipt.shape_id)
            .unwrap();
        assert_eq!(pasted.children.len(), 2);
        let pasted_connector = pasted
            .children
            .iter()
            .find(|child| child.copy_source_id == Some(1))
            .unwrap();
        let pasted_target = pasted
            .children
            .iter()
            .find(|child| child.copy_source_id == Some(11))
            .unwrap();
        let package = session.package().unwrap();
        let part = package.page_part_paths[0].clone();
        let carried = package.page_contents[&part]
            .connects()
            .filter(|connect| connect.from_sheet == pasted_connector.source_id)
            .collect::<Vec<_>>();
        assert_eq!(carried.len(), 1);
        assert_eq!(carried[0].from_cell.as_deref(), Some("BeginX"));
        assert_eq!(carried[0].to_sheet, pasted_target.source_id);
        assert_eq!(carried[0].to_cell.as_deref(), Some("Connections.X1"));
        assert_eq!(
            package.page_contents[&part]
                .connects()
                .filter(|connect| connect.from_sheet == 1)
                .count(),
            6
        );
    }

    #[test]
    fn pasted_shapes_keep_their_source_text_tokens() {
        let session = DiagramSession::open(
            include_bytes!("../../vsdx-parse/tests/fixtures/text-accounting.vsdx"),
            915,
        )
        .unwrap();
        let snapshot = session.snapshot().unwrap();
        let edited = find_by_source(&snapshot.pages[0].shapes, 3).id.clone();
        session
            .set_shape_text(&EditCtx::local("text"), "page:1", &edited, "edited")
            .unwrap();
        let mut pasted_sources = Vec::new();
        for source_id in [1, 2, 3] {
            let snapshot = session.snapshot().unwrap();
            let source = find_by_source(&snapshot.pages[0].shapes, source_id).clone();
            let draft = tree_draft(&session, "page:1", &source);
            let receipt = session
                .add_shape_tree(&EditCtx::local("paste"), "page:1", &draft)
                .unwrap();
            let resnapshot = session.snapshot().unwrap();
            pasted_sources.push(
                resnapshot.pages[0]
                    .shapes
                    .iter()
                    .find(|shape| shape.id == receipt.shape_id)
                    .unwrap()
                    .source_id,
            );
        }
        let package = session.package().unwrap();
        let part = package.page_part_paths[0].clone();
        let text_of = |source_id: u32| {
            package.page_contents[&part]
                .shapes()
                .find(|shape| shape.id == source_id)
                .unwrap()
                .text()
                .unwrap()
                .to_vec()
        };
        assert_eq!(
            text_of(pasted_sources[1]),
            vec![vsdx_parse::TextToken::Field(0)]
        );
        assert_eq!(
            text_of(pasted_sources[0]),
            vec![
                vsdx_parse::TextToken::CharacterRun(0),
                vsdx_parse::TextToken::ParagraphRun(0),
            ]
        );
        assert_eq!(
            text_of(pasted_sources[2]),
            vec![vsdx_parse::TextToken::Literal("edited".to_owned())]
        );
        let expected = pasted_sources
            .iter()
            .map(|source_id| text_of(*source_id))
            .collect::<Vec<_>>();
        let saved = session.save().unwrap();
        let reopened = DiagramSession::open(&saved, 916).unwrap();
        assert_reopened_projection_eq(&session, &reopened, "token-preserving paste");
        let reopened_package = reopened.package().unwrap();
        for (source_id, tokens) in pasted_sources.iter().zip(expected) {
            let actual = reopened_package.page_contents[&part]
                .shapes()
                .find(|shape| shape.id == *source_id)
                .unwrap()
                .text()
                .unwrap()
                .to_vec();
            assert_eq!(actual, tokens);
        }
    }

    #[test]
    fn pasted_shapes_remember_their_source_page() {
        let session = grouped_glue_session();
        let (pasted, _) = paste_group(&session, "page:1", 10);
        assert_eq!(pasted.copy_source_id, Some(10));
        assert_eq!(pasted.copy_source_page_id, Some(1));
        assert_eq!(pasted.children[0].copy_source_id, Some(11));
        assert_eq!(pasted.children[0].copy_source_page_id, Some(1));
    }

    #[test]
    fn cut_paste_without_its_source_preserves_text_tokens() {
        let session = DiagramSession::open(
            include_bytes!("../../vsdx-parse/tests/fixtures/text-accounting.vsdx"),
            917,
        )
        .unwrap();
        let snapshot = session.snapshot().unwrap();
        let source = find_by_source(&snapshot.pages[0].shapes, 2).clone();
        let draft = tree_draft(&session, "page:1", &source);
        assert_eq!(draft.copy_source_page_id, Some(1));
        session
            .delete_shape(&EditCtx::local("cut"), "page:1", &source.id)
            .unwrap();
        let receipt = session
            .add_shape_tree(&EditCtx::local("paste"), "page:1", &draft)
            .unwrap();
        let resnapshot = session.snapshot().unwrap();
        let pasted = resnapshot.pages[0]
            .shapes
            .iter()
            .find(|shape| shape.id == receipt.shape_id)
            .unwrap();
        assert_eq!(pasted.copy_source_page_id, Some(1));
        let package = session.package().unwrap();
        let part = package.page_part_paths[0].clone();
        let tokens = package.page_contents[&part]
            .shapes()
            .find(|shape| shape.id == pasted.source_id)
            .unwrap()
            .text()
            .unwrap()
            .to_vec();
        assert_eq!(tokens, vec![vsdx_parse::TextToken::Field(0)]);
        let reopened = DiagramSession::open(&session.save().unwrap(), 918).unwrap();
        assert_reopened_projection_eq(&session, &reopened, "cut paste preserves field");
    }

    #[test]
    fn detached_paste_with_an_unknown_source_page_is_refused() {
        let session = grouped_glue_session();
        let snapshot = session.snapshot().unwrap();
        let source = find_by_source(&snapshot.pages[0].shapes, 10).clone();
        let mut draft = tree_draft(&session, "page:1", &source);
        draft.copy_source_page_id = Some(99);
        draft.children[0].copy_source_page_id = Some(99);
        session
            .delete_shape(&EditCtx::local("cut"), "page:1", &source.id)
            .unwrap();
        let before = session.encode_state_as_update_v1();
        let error = session
            .add_shape_tree(&EditCtx::local("paste"), "page:1", &draft)
            .unwrap_err();
        assert!(
            error.to_string().contains("paste source page is missing"),
            "unexpected refusal: {error}"
        );
        assert_eq!(session.encode_state_as_update_v1(), before);
    }

    #[test]
    fn detached_paste_spanning_source_pages_is_refused() {
        let session = grouped_glue_session();
        let snapshot = session.snapshot().unwrap();
        let source = find_by_source(&snapshot.pages[0].shapes, 10).clone();
        let mut draft = tree_draft(&session, "page:1", &source);
        draft.children[0].copy_source_page_id = Some(2);
        session
            .delete_shape(&EditCtx::local("cut"), "page:1", &source.id)
            .unwrap();
        let before = session.encode_state_as_update_v1();
        let error = session
            .add_shape_tree(&EditCtx::local("paste"), "page:1", &draft)
            .unwrap_err();
        assert!(
            error.to_string().contains("span multiple pages"),
            "unexpected refusal: {error}"
        );
        assert_eq!(session.encode_state_as_update_v1(), before);
    }

    #[test]
    fn copies_of_copies_stay_self_referential() {
        for save_between in [false, true] {
            let session = grouped_glue_session();
            let snapshot = session.snapshot().unwrap();
            let child_id = find_by_source(&snapshot.pages[0].shapes, 11).id.clone();
            session
                .set_cell_formula(
                    &EditCtx::local("edit"),
                    "page:1",
                    &child_id,
                    "PinX",
                    "Sheet.10!Width/2",
                )
                .unwrap();
            let (first, _) = paste_group(&session, "page:1", 10);
            let session = if save_between {
                DiagramSession::open(&session.save().unwrap(), 914).unwrap()
            } else {
                session
            };
            let snapshot = session.snapshot().unwrap();
            let source = snapshot.pages[0]
                .shapes
                .iter()
                .find(|shape| shape.source_id == first.source_id)
                .cloned()
                .unwrap();
            let draft = tree_draft(&session, "page:1", &source);
            let receipt = session
                .add_shape_tree(&EditCtx::local("paste"), "page:1", &draft)
                .unwrap();
            let snapshot = session.snapshot().unwrap();
            let second = snapshot.pages[0]
                .shapes
                .iter()
                .find(|shape| shape.id == receipt.shape_id)
                .unwrap();
            let expected = format!("Sheet.{}!Width/2", second.source_id);
            assert_eq!(
                cell_formula(&second.children[0], "PinX"),
                Some(expected.as_str()),
                "save_between={save_between}"
            );
            let saved = session.save().unwrap();
            let reopened = DiagramSession::open(&saved, 915).unwrap();
            assert_reopened_projection_eq(&session, &reopened, "second-generation paste");
            let resnapshot = reopened.snapshot().unwrap();
            let root = find_by_source(&resnapshot.pages[0].shapes, second.source_id);
            assert_eq!(
                cell_formula(&root.children[0], "PinX"),
                Some(expected.as_str()),
                "save_between={save_between}"
            );
        }
    }
}
