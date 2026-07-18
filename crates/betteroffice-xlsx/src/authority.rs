use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::sync::Arc;

use sha2::{Digest, Sha256};
use xlsx_model::{
    Cell, CellRange, CellRef, CellValue, DateSystem, ErrorValue, MAX_COLS, MAX_ROWS, Sheet,
    SheetId, Stylesheet, Workbook as WorkbookModel,
};
use xlsx_ops::Op;
use yrs::updates::decoder::{Decode, Decoder, DecoderV1};
use yrs::updates::encoder::Encode;
use yrs::{
    Any, Array, ArrayRef, BranchID, Doc, Map, MapPrelim, MapRef, Out, ReadTxn, StateVector,
    Subscription, Transact, TransactionMut, Update, WriteTxn,
};

const META: &str = "xlsx";
const SHEET_ORDER: &str = "xlsx:sheet-order";
const SHEETS: &str = "xlsx:sheets";
const SCHEMA_VERSION: i64 = 2;
const BASE_FINGERPRINT: &str = "baseFingerprint";
const STRUCTURE_GENERATION: &str = "structureGeneration";
const CONTENTS: &str = "contents";
const COL_WIDTHS: &str = "colWidths";
const MERGES: &str = "merges";
const NAME: &str = "name";
const ROW_HEIGHTS: &str = "rowHeights";
const STYLES: &str = "styles";
const BOOTSTRAP_ORIGIN: &str = "xlsx:bootstrap";
const HYDRATE_ORIGIN: &str = "xlsx:hydrate";
const REMOTE_ORIGIN: &str = "xlsx:remote";
const MAX_SAFE_CLIENT_ID: u64 = (1_u64 << 53) - 1;
pub(crate) const MAX_STATE_VECTOR_ENTRIES: u32 = 65_536;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum SyncOrigin {
    User,
    Agent,
    Undo,
    Redo,
}

impl SyncOrigin {
    fn as_str(self) -> &'static str {
        match self {
            Self::User => "xlsx:user",
            Self::Agent => "xlsx:agent",
            Self::Undo => "xlsx:undo",
            Self::Redo => "xlsx:redo",
        }
    }
}

#[derive(Debug)]
pub(crate) enum AuthorityError {
    ClientIdConflict(u64),
    InvalidStateVector(String),
    InvalidUpdate(String),
    InvalidState(String),
    Observer(String),
}

#[derive(Clone)]
struct WorkbookBase {
    bootstrap_client_id: u64,
    date_system: DateSystem,
    fingerprint: String,
    shared_strings: Vec<String>,
    styles: Stylesheet,
}

impl WorkbookBase {
    fn from_model(model: &WorkbookModel) -> Result<Self, String> {
        let (fingerprint, bootstrap_client_id) = fingerprint_model(model)?;
        Ok(Self {
            bootstrap_client_id,
            date_system: model.date_system,
            fingerprint,
            shared_strings: model.shared_strings.clone(),
            styles: model.styles.clone(),
        })
    }

    fn workbook(&self) -> WorkbookModel {
        WorkbookModel {
            sheets: Vec::new(),
            date_system: self.date_system,
            shared_strings: self.shared_strings.clone(),
            styles: self.styles.clone(),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct WorkbookStructure {
    generation: i64,
    sheet_keys: Vec<String>,
    sheet_names: Vec<String>,
    merges: Vec<Vec<CellRange>>,
    shared_types: BTreeMap<String, SheetSharedTypes>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SheetSharedTypes {
    sheet: BranchID,
    col_widths: BranchID,
    contents: BranchID,
    row_heights: BranchID,
    styles: BranchID,
}

pub(crate) struct StagedUpdate {
    pub(crate) effective: bool,
    pub(crate) model: WorkbookModel,
    pub(crate) pending: bool,
    pub(crate) structure: WorkbookStructure,
    pub(crate) update: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SheetOrderEntry {
    before: Vec<String>,
    after: Vec<String>,
}

#[derive(Default)]
struct SheetOrderHistory {
    undo: Vec<SheetOrderEntry>,
    redo: Vec<SheetOrderEntry>,
}

enum HistoryAction {
    Push(SheetOrderEntry),
    Undo(SheetOrderEntry),
    Redo(SheetOrderEntry),
}

pub(crate) struct WorkbookAuthority {
    doc: Doc,
    base: WorkbookBase,
    history: SheetOrderHistory,
    next_sheet_id: u64,
}

impl WorkbookAuthority {
    pub(crate) fn from_model(model: &WorkbookModel) -> Result<Self, AuthorityError> {
        Self::from_model_internal(model, None)
    }

    pub(crate) fn from_model_with_client_id(
        model: &WorkbookModel,
        client_id: u64,
    ) -> Result<Self, AuthorityError> {
        Self::from_model_internal(model, Some(client_id))
    }

    fn from_model_internal(
        model: &WorkbookModel,
        client_id: Option<u64>,
    ) -> Result<Self, AuthorityError> {
        let base = WorkbookBase::from_model(model).map_err(AuthorityError::InvalidState)?;
        if client_id == Some(base.bootstrap_client_id) {
            return Err(AuthorityError::ClientIdConflict(base.bootstrap_client_id));
        }

        let bootstrap = Doc::with_client_id(base.bootstrap_client_id);
        let keys = (0..model.sheets.len())
            .map(|index| format!("sheet:{index}"))
            .collect::<Vec<_>>();
        seed(&bootstrap, &base, model, &keys);
        let bootstrap_update = bootstrap
            .transact()
            .encode_state_as_update_v1(&StateVector::default());

        let doc = match client_id {
            Some(client_id) => Doc::with_client_id(client_id),
            None => loop {
                let candidate = Doc::new();
                if candidate.client_id().get() != base.bootstrap_client_id {
                    break candidate;
                }
            },
        };
        hydrate_doc(&doc, &bootstrap_update).map_err(AuthorityError::InvalidState)?;
        let authority = Self {
            doc,
            base,
            history: SheetOrderHistory::default(),
            next_sheet_id: 0,
        };
        authority
            .strict_materialize()
            .map_err(AuthorityError::InvalidState)?;
        Ok(authority)
    }

    pub(crate) fn client_id(&self) -> u64 {
        self.doc.client_id().get()
    }

    pub(crate) fn materialize(&self) -> Result<WorkbookModel, AuthorityError> {
        self.materialize_internal(false)
            .map(|(model, _)| model)
            .map_err(AuthorityError::InvalidState)
    }

    pub(crate) fn structure(&self) -> Result<WorkbookStructure, AuthorityError> {
        self.materialize_internal(false)
            .map(|(_, structure)| structure)
            .map_err(AuthorityError::InvalidState)
    }

    pub(crate) fn apply_ops(
        &mut self,
        ops: &[Op],
        origin: SyncOrigin,
    ) -> Result<(), AuthorityError> {
        let mut model = self.materialize()?;
        for op in ops {
            xlsx_ops::apply(&mut model, op).map_err(|error| {
                AuthorityError::InvalidState(format!(
                    "cannot apply local operation to authored state: {error}"
                ))
            })?;
        }
        self.sync_model(&model, ops, origin)
            .map_err(AuthorityError::InvalidState)
    }

    pub(crate) fn encode_state_vector_v1(&self) -> Vec<u8> {
        self.doc.transact().state_vector().encode_v1()
    }

    pub(crate) fn encode_state_as_update_v1(&self) -> Vec<u8> {
        self.doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default())
    }

    pub(crate) fn encode_diff_v1(
        &self,
        remote_state_vector: &[u8],
    ) -> Result<Vec<u8>, AuthorityError> {
        let state_vector = decode_state_vector_v1(remote_state_vector)
            .map_err(AuthorityError::InvalidStateVector)?;
        Ok(self.doc.transact().encode_diff_v1(&state_vector))
    }

    pub(crate) fn stage_updates_v1(
        &self,
        updates: &[&[u8]],
    ) -> Result<StagedUpdate, AuthorityError> {
        if updates.is_empty() {
            return Err(AuthorityError::InvalidUpdate(
                "no updates were provided".to_string(),
            ));
        }
        let decoded = updates
            .iter()
            .map(|update| decode_update_v1(update).map_err(AuthorityError::InvalidUpdate))
            .collect::<Result<Vec<_>, _>>()?;
        let (incoming, live_update) = if updates.len() == 1 {
            (decoded.into_iter().next().unwrap(), updates[0].to_vec())
        } else {
            let update = Update::merge_updates(decoded);
            let encoded = update.encode_v1();
            (update, encoded)
        };
        let before = self.encode_state_as_update_v1();
        let staged_doc = Doc::with_client_id(self.client_id());
        hydrate_doc(&staged_doc, &before).map_err(AuthorityError::InvalidState)?;
        staged_doc
            .transact_mut_with(REMOTE_ORIGIN)
            .apply_update(incoming)
            .map_err(|error| AuthorityError::InvalidUpdate(error.to_string()))?;

        let staged = Self {
            doc: staged_doc,
            base: self.base.clone(),
            history: SheetOrderHistory::default(),
            next_sheet_id: self.next_sheet_id,
        };
        let after = staged.encode_state_as_update_v1();
        let pending = {
            let txn = staged.doc.transact();
            txn.store().pending_update().is_some() || txn.store().pending_ds().is_some()
        };
        let (model, structure) = staged
            .strict_materialize()
            .map_err(AuthorityError::InvalidState)?;
        Ok(StagedUpdate {
            effective: before != after,
            model,
            pending,
            structure,
            update: live_update,
        })
    }

    pub(crate) fn apply_update_v1(&self, update: &[u8]) -> Result<(), AuthorityError> {
        let update = decode_update_v1(update).map_err(AuthorityError::InvalidUpdate)?;
        self.doc
            .transact_mut_with(REMOTE_ORIGIN)
            .apply_update(update)
            .map_err(|error| AuthorityError::InvalidUpdate(error.to_string()))
    }

    pub(crate) fn observe_update_v1<F>(&self, callback: F) -> Result<Subscription, AuthorityError>
    where
        F: Fn(bool, Vec<u8>) + 'static,
    {
        self.doc
            .observe_update_v1(move |txn, event| {
                let remote = txn
                    .origin()
                    .is_some_and(|origin| origin.as_ref() == REMOTE_ORIGIN.as_bytes());
                let _ = catch_unwind(AssertUnwindSafe(|| {
                    callback(remote, event.update.clone());
                }));
            })
            .map_err(|error| AuthorityError::Observer(error.to_string()))
    }

    pub(crate) fn clear_history(&mut self) {
        self.history = SheetOrderHistory::default();
    }

    fn strict_materialize(&self) -> Result<(WorkbookModel, WorkbookStructure), String> {
        self.materialize_internal(true)
    }

    fn materialize_internal(
        &self,
        strict: bool,
    ) -> Result<(WorkbookModel, WorkbookStructure), String> {
        let txn = self.doc.transact();
        if strict {
            require_root_keys(&txn, &[META, SHEET_ORDER, SHEETS])?;
        }
        let meta = txn
            .get_map(META)
            .ok_or_else(|| "missing workbook metadata".to_string())?;
        if strict {
            require_map_keys(
                &meta,
                &txn,
                &[BASE_FINGERPRINT, "schemaVersion", STRUCTURE_GENERATION],
                "workbook metadata",
            )?;
        }
        let version = meta
            .get(&txn, "schemaVersion")
            .and_then(|value| value.cast::<i64>().ok())
            .ok_or_else(|| "missing schema version".to_string())?;
        if version != SCHEMA_VERSION {
            return Err(format!("unsupported schema version {version}"));
        }
        let fingerprint = meta
            .get(&txn, BASE_FINGERPRINT)
            .and_then(|value| value.cast::<String>().ok())
            .ok_or_else(|| "missing workbook base fingerprint".to_string())?;
        if fingerprint != self.base.fingerprint {
            return Err("workbook base fingerprint does not match shared state".to_string());
        }
        let generation = structure_generation(&meta, &txn)?;

        let order = txn
            .get_array(SHEET_ORDER)
            .ok_or_else(|| "missing sheet order".to_string())?;
        let sheets = txn
            .get_map(SHEETS)
            .ok_or_else(|| "missing sheet map".to_string())?;
        let keys = sheet_keys(&order, &txn)?;
        let mut seen = HashSet::with_capacity(keys.len());
        let mut model = self.base.workbook();
        for key in &keys {
            if !seen.insert(key.clone()) {
                return Err(format!("duplicate sheet key {key}"));
            }
            let sheet_map = sheets
                .get(&txn, key)
                .and_then(|value| value.cast::<MapRef>().ok())
                .ok_or_else(|| format!("missing sheet {key}"))?;
            if strict {
                require_map_keys(
                    &sheet_map,
                    &txn,
                    &[COL_WIDTHS, CONTENTS, MERGES, NAME, ROW_HEIGHTS, STYLES],
                    &format!("sheet {key}"),
                )?;
            }
            model.sheets.push(materialize_sheet(&sheet_map, &txn)?);
        }
        let active = keys.iter().cloned().collect::<BTreeSet<_>>();
        let all_keys = sheets
            .keys(&txn)
            .map(str::to_string)
            .collect::<BTreeSet<_>>();
        let mut shared_types = BTreeMap::new();
        for key in all_keys {
            let sheet_map = sheets
                .get(&txn, &key)
                .and_then(|value| value.cast::<MapRef>().ok())
                .ok_or_else(|| format!("sheet {key} is not a map"))?;
            if strict && !active.contains(&key) {
                require_map_keys(
                    &sheet_map,
                    &txn,
                    &[COL_WIDTHS, CONTENTS, MERGES, NAME, ROW_HEIGHTS, STYLES],
                    &format!("inactive sheet {key}"),
                )?;
                materialize_sheet(&sheet_map, &txn)?;
            }
            shared_types.insert(key, sheet_shared_types(&sheet_map, &txn)?);
        }
        let structure = WorkbookStructure {
            generation,
            sheet_keys: keys,
            sheet_names: model
                .sheets
                .iter()
                .map(|sheet| sheet.name.clone())
                .collect(),
            merges: model
                .sheets
                .iter()
                .map(|sheet| sheet.merges.clone())
                .collect(),
            shared_types,
        };
        Ok((model, structure))
    }

    fn sync_model(
        &mut self,
        model: &WorkbookModel,
        ops: &[Op],
        origin: SyncOrigin,
    ) -> Result<(), String> {
        let current_keys = self.current_sheet_keys()?;
        let (keys, history) =
            self.plan_sheet_keys(&current_keys, ops, model.sheets.len(), origin)?;
        self.validate_sync_state(&current_keys, &keys)?;

        let topology_changed = current_keys != keys;
        let full_sync = ops.iter().any(requires_full_semantic_sync);
        let structure_delta = i64::try_from(ops.iter().filter(|op| is_structural_op(op)).count())
            .map_err(|_| "too many structural operations".to_string())?;
        let mut authored_cells = HashSet::new();
        let mut col_widths = HashSet::new();
        let mut row_heights = HashSet::new();
        let mut merges = HashSet::new();
        if !full_sync {
            let targets = targeted_sheet_keys(&current_keys, &keys, ops)?;
            for (op, target) in ops.iter().zip(targets) {
                match (op, target) {
                    (Op::SetCell { at, .. }, Some(key)) => {
                        authored_cells.insert((key, *at));
                    }
                    (Op::SetColWidth { col, .. }, Some(key)) => {
                        col_widths.insert((key, *col));
                    }
                    (Op::SetRowHeight { row, .. }, Some(key)) => {
                        row_heights.insert((key, *row));
                    }
                    (Op::MergeCells { .. } | Op::UnmergeCells { .. }, Some(key)) => {
                        merges.insert(key);
                    }
                    (Op::AddSheet { .. } | Op::RemoveSheet { .. }, None) => {}
                    (_, None) => {}
                    _ => return Err("semantic operation requires a full sync".to_string()),
                }
            }
        }
        let current_key_set = current_keys
            .iter()
            .map(String::as_str)
            .collect::<HashSet<_>>();
        let newly_active = keys
            .iter()
            .filter(|key| !current_key_set.contains(key.as_str()))
            .cloned()
            .collect::<HashSet<_>>();
        if !topology_changed
            && !full_sync
            && structure_delta == 0
            && authored_cells.is_empty()
            && col_widths.is_empty()
            && row_heights.is_empty()
            && merges.is_empty()
        {
            self.apply_history(history);
            return Ok(());
        }

        let mut txn = self.doc.transact_mut_with(origin.as_str());
        let order = txn
            .get_array(SHEET_ORDER)
            .ok_or_else(|| "missing sheet order".to_string())?;
        let sheets = txn
            .get_map(SHEETS)
            .ok_or_else(|| "missing sheet map".to_string())?;
        if structure_delta != 0 {
            let meta = txn
                .get_map(META)
                .ok_or_else(|| "missing workbook metadata".to_string())?;
            let next = structure_generation(&meta, &txn)?
                .checked_add(structure_delta)
                .ok_or_else(|| "structure generation overflow".to_string())?;
            meta.try_update(&mut txn, STRUCTURE_GENERATION, next);
        }
        if topology_changed {
            patch_sheet_order(&order, &mut txn, &current_keys, &keys)?;
        }
        if full_sync {
            for (key, sheet) in keys.iter().zip(&model.sheets) {
                let sheet_map = sheet_map_for_sync(&sheets, &mut txn, key)?;
                sync_sheet(&sheet_map, &mut txn, sheet);
            }
        } else {
            for (key, sheet) in keys.iter().zip(&model.sheets) {
                if newly_active.contains(key) {
                    let sheet_map = sheet_map_for_sync(&sheets, &mut txn, key)?;
                    sync_sheet(&sheet_map, &mut txn, sheet);
                }
            }
            for (key, at) in authored_cells {
                let (sheet_map, sheet_model) =
                    sheet_parts_by_key(&sheets, &txn, &keys, model, &key)?;
                sync_authored_cell(&sheet_map, &mut txn, sheet_model, at);
            }
            for (key, col) in col_widths {
                let (sheet_map, sheet_model) =
                    sheet_parts_by_key(&sheets, &txn, &keys, model, &key)?;
                let map: MapRef = sheet_map.get_or_init(&mut txn, COL_WIDTHS);
                sync_number(
                    &map,
                    &mut txn,
                    col,
                    sheet_model.col_widths.get(&col).copied(),
                );
            }
            for (key, row) in row_heights {
                let (sheet_map, sheet_model) =
                    sheet_parts_by_key(&sheets, &txn, &keys, model, &key)?;
                let map: MapRef = sheet_map.get_or_init(&mut txn, ROW_HEIGHTS);
                sync_number(
                    &map,
                    &mut txn,
                    row,
                    sheet_model.row_heights.get(&row).copied(),
                );
            }
            for key in merges {
                let (sheet_map, sheet_model) =
                    sheet_parts_by_key(&sheets, &txn, &keys, model, &key)?;
                sheet_map.try_update(&mut txn, MERGES, merges_to_any(&sheet_model.merges));
            }
        }
        drop(txn);
        self.apply_history(history);
        Ok(())
    }

    fn current_sheet_keys(&self) -> Result<Vec<String>, String> {
        let txn = self.doc.transact();
        let order = txn
            .get_array(SHEET_ORDER)
            .ok_or_else(|| "missing sheet order".to_string())?;
        sheet_keys(&order, &txn)
    }

    fn validate_sync_state(&self, current: &[String], desired: &[String]) -> Result<(), String> {
        let txn = self.doc.transact();
        let sheets = txn
            .get_map(SHEETS)
            .ok_or_else(|| "missing sheet map".to_string())?;
        for key in current {
            match sheets.get(&txn, key) {
                Some(Out::YMap(_)) => {}
                Some(_) => return Err(format!("sheet {key} is not a map")),
                None => return Err(format!("missing sheet {key}")),
            }
        }
        for key in desired {
            if let Some(value) = sheets.get(&txn, key)
                && !matches!(value, Out::YMap(_))
            {
                return Err(format!("sheet {key} is not a map"));
            }
        }
        Ok(())
    }

    fn plan_sheet_keys(
        &mut self,
        current: &[String],
        ops: &[Op],
        final_len: usize,
        origin: SyncOrigin,
    ) -> Result<(Vec<String>, HistoryAction), String> {
        match origin {
            SyncOrigin::User | SyncOrigin::Agent => {
                let keys = self.reconcile_sheet_keys(current.to_vec(), ops, final_len)?;
                let entry = SheetOrderEntry {
                    before: current.to_vec(),
                    after: keys.clone(),
                };
                Ok((keys, HistoryAction::Push(entry)))
            }
            SyncOrigin::Undo => {
                let entry = self
                    .history
                    .undo
                    .last()
                    .cloned()
                    .ok_or_else(|| "sheet-order undo history is empty".to_string())?;
                if entry.after != current {
                    return Err("sheet-order undo history does not match current state".to_string());
                }
                if entry.before.len() != final_len {
                    return Err("sheet-order undo result does not match workbook".to_string());
                }
                Ok((entry.before.clone(), HistoryAction::Undo(entry)))
            }
            SyncOrigin::Redo => {
                let entry = self
                    .history
                    .redo
                    .last()
                    .cloned()
                    .ok_or_else(|| "sheet-order redo history is empty".to_string())?;
                if entry.before != current {
                    return Err("sheet-order redo history does not match current state".to_string());
                }
                if entry.after.len() != final_len {
                    return Err("sheet-order redo result does not match workbook".to_string());
                }
                Ok((entry.after.clone(), HistoryAction::Redo(entry)))
            }
        }
    }

    fn reconcile_sheet_keys(
        &mut self,
        mut keys: Vec<String>,
        ops: &[Op],
        final_len: usize,
    ) -> Result<Vec<String>, String> {
        for op in ops {
            match op {
                Op::AddSheet { index, .. } => {
                    if *index > keys.len() {
                        return Err(format!("sheet insertion index {index} is out of range"));
                    }
                    let key = self.allocate_sheet_key();
                    keys.insert(*index, key);
                }
                Op::RemoveSheet { index } => {
                    if *index >= keys.len() {
                        return Err(format!("sheet removal index {index} is out of range"));
                    }
                    keys.remove(*index);
                }
                _ => {}
            }
        }
        if keys.len() != final_len {
            return Err("sheet order does not match workbook projection".to_string());
        }
        Ok(keys)
    }

    fn apply_history(&mut self, action: HistoryAction) {
        match action {
            HistoryAction::Push(entry) => {
                self.history.undo.push(entry);
                self.history.redo.clear();
            }
            HistoryAction::Undo(entry) => {
                self.history.undo.pop();
                self.history.redo.push(entry);
            }
            HistoryAction::Redo(entry) => {
                self.history.redo.pop();
                self.history.undo.push(entry);
            }
        }
    }

    fn allocate_sheet_key(&mut self) -> String {
        let key = format!("replica:{}:{}", self.client_id(), self.next_sheet_id);
        self.next_sheet_id += 1;
        key
    }
}

pub(crate) fn is_structural_op(op: &Op) -> bool {
    matches!(
        op,
        Op::InsertRows { .. }
            | Op::DeleteRows { .. }
            | Op::InsertCols { .. }
            | Op::DeleteCols { .. }
            | Op::MergeCells { .. }
            | Op::UnmergeCells { .. }
            | Op::AddSheet { .. }
            | Op::RemoveSheet { .. }
            | Op::RenameSheet { .. }
            | Op::RestoreSheet { .. }
    )
}

fn seed(doc: &Doc, base: &WorkbookBase, model: &WorkbookModel, keys: &[String]) {
    let mut txn = doc.transact_mut_with(BOOTSTRAP_ORIGIN);
    let meta = txn.get_or_insert_map(META);
    meta.insert(&mut txn, BASE_FINGERPRINT, base.fingerprint.as_str());
    meta.insert(&mut txn, "schemaVersion", SCHEMA_VERSION);
    meta.insert(&mut txn, STRUCTURE_GENERATION, 0_i64);
    let order = txn.get_or_insert_array(SHEET_ORDER);
    order.insert_range(&mut txn, 0, keys.iter().cloned());
    let sheets = txn.get_or_insert_map(SHEETS);
    for (key, sheet) in keys.iter().zip(&model.sheets) {
        let sheet_map = sheets.insert(&mut txn, key.as_str(), MapPrelim::default());
        sync_sheet(&sheet_map, &mut txn, sheet);
    }
}

fn hydrate_doc(doc: &Doc, update: &[u8]) -> Result<(), String> {
    let update = decode_update_v1(update)?;
    doc.transact_mut_with(HYDRATE_ORIGIN)
        .apply_update(update)
        .map_err(|error| error.to_string())
}

fn decode_state_vector_v1(bytes: &[u8]) -> Result<StateVector, String> {
    validate_state_vector_entry_count(bytes)?;
    let mut decoder = DecoderV1::from(bytes);
    let state_vector = StateVector::decode(&mut decoder).map_err(|error| error.to_string())?;
    if !decoder
        .read_to_end()
        .map_err(|error| error.to_string())?
        .is_empty()
    {
        return Err("state vector contains trailing bytes".to_string());
    }
    Ok(state_vector)
}

fn decode_update_v1(bytes: &[u8]) -> Result<Update, String> {
    let mut decoder = DecoderV1::from(bytes);
    let update = Update::decode(&mut decoder).map_err(|error| error.to_string())?;
    if !decoder
        .read_to_end()
        .map_err(|error| error.to_string())?
        .is_empty()
    {
        return Err("update contains trailing bytes".to_string());
    }
    Ok(update)
}

fn validate_state_vector_entry_count(bytes: &[u8]) -> Result<(), String> {
    let Some((&first, _)) = bytes.split_first() else {
        return Err("state vector is empty".to_string());
    };
    let mut value = u32::from(first & 0x7f);
    let mut shift = 7;
    let mut used = 1;
    let mut byte = first;
    while byte & 0x80 != 0 {
        if used == 5 || used >= bytes.len() {
            return Err("invalid state vector entry count".to_string());
        }
        byte = bytes[used];
        if used == 4 && byte > 0x0f {
            return Err("invalid state vector entry count".to_string());
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
        return Err("state vector entry count exceeds its payload".to_string());
    }
    Ok(())
}

fn requires_full_semantic_sync(op: &Op) -> bool {
    matches!(
        op,
        Op::InsertRows { .. }
            | Op::DeleteRows { .. }
            | Op::InsertCols { .. }
            | Op::DeleteCols { .. }
            | Op::RenameSheet { .. }
            | Op::RestoreSheet { .. }
    )
}

#[derive(Clone)]
enum SheetToken {
    Existing(String),
    Added(usize),
}

fn targeted_sheet_keys(
    current: &[String],
    desired: &[String],
    ops: &[Op],
) -> Result<Vec<Option<String>>, String> {
    let mut tokens = current
        .iter()
        .cloned()
        .map(SheetToken::Existing)
        .collect::<Vec<_>>();
    let mut targets = Vec::with_capacity(ops.len());
    let mut next_added = 0;
    for op in ops {
        match op {
            Op::AddSheet { index, .. } => {
                if *index > tokens.len() {
                    return Err(format!("sheet insertion index {index} is out of range"));
                }
                tokens.insert(*index, SheetToken::Added(next_added));
                next_added += 1;
                targets.push(None);
            }
            Op::RemoveSheet { index } => {
                if *index >= tokens.len() {
                    return Err(format!("sheet removal index {index} is out of range"));
                }
                tokens.remove(*index);
                targets.push(None);
            }
            op => {
                let sheet = op_sheet(op)
                    .ok_or_else(|| "operation has no sheet target".to_string())?
                    .0 as usize;
                let token = tokens
                    .get(sheet)
                    .cloned()
                    .ok_or_else(|| format!("sheet {sheet} is out of range"))?;
                targets.push(Some(token));
            }
        }
    }
    if tokens.len() != desired.len() {
        return Err("sheet operation plan does not match final order".to_string());
    }
    let mut added = HashMap::new();
    for (token, key) in tokens.iter().zip(desired) {
        match token {
            SheetToken::Existing(existing) if existing != key => {
                return Err("sheet operation plan changed retained identity".to_string());
            }
            SheetToken::Existing(_) => {}
            SheetToken::Added(id) => {
                added.insert(*id, key.clone());
            }
        }
    }
    let active = desired.iter().map(String::as_str).collect::<HashSet<_>>();
    Ok(targets
        .into_iter()
        .map(|target| match target {
            Some(SheetToken::Existing(key)) if active.contains(key.as_str()) => Some(key),
            Some(SheetToken::Added(id)) => added.get(&id).cloned(),
            _ => None,
        })
        .collect())
}

fn op_sheet(op: &Op) -> Option<SheetId> {
    match op {
        Op::SetCell { sheet, .. }
        | Op::InsertRows { sheet, .. }
        | Op::DeleteRows { sheet, .. }
        | Op::InsertCols { sheet, .. }
        | Op::DeleteCols { sheet, .. }
        | Op::SetColWidth { sheet, .. }
        | Op::SetRowHeight { sheet, .. }
        | Op::MergeCells { sheet, .. }
        | Op::UnmergeCells { sheet, .. }
        | Op::RenameSheet { sheet, .. }
        | Op::RestoreSheet { sheet, .. } => Some(*sheet),
        Op::AddSheet { .. } | Op::RemoveSheet { .. } => None,
    }
}

fn patch_sheet_order(
    order: &ArrayRef,
    txn: &mut TransactionMut<'_>,
    existing: &[String],
    desired: &[String],
) -> Result<(), String> {
    let mut working = existing.to_vec();
    let mut index = 0;
    while index < desired.len() {
        if working.get(index) == desired.get(index) {
            index += 1;
            continue;
        }
        if let Some(offset) = working[index..]
            .iter()
            .position(|key| key == &desired[index])
        {
            order.remove_range(txn, yrs_index(index)?, yrs_index(offset)?);
            working.drain(index..index + offset);
        } else {
            order.insert(txn, yrs_index(index)?, desired[index].clone());
            working.insert(index, desired[index].clone());
            index += 1;
        }
    }
    if working.len() > desired.len() {
        order.remove_range(
            txn,
            yrs_index(desired.len())?,
            yrs_index(working.len() - desired.len())?,
        );
    }
    Ok(())
}

fn yrs_index(index: usize) -> Result<u32, String> {
    u32::try_from(index).map_err(|_| "sheet order exceeds Yrs index range".to_string())
}

fn sheet_map_for_sync(
    sheets: &MapRef,
    txn: &mut TransactionMut<'_>,
    key: &str,
) -> Result<MapRef, String> {
    match sheets.get(txn, key) {
        Some(Out::YMap(map)) => Ok(map),
        Some(_) => Err(format!("sheet {key} is not a map")),
        None => Ok(sheets.insert(txn, key, MapPrelim::default())),
    }
}

fn sheet_parts_by_key<'a, T: ReadTxn>(
    sheets: &MapRef,
    txn: &T,
    keys: &[String],
    model: &'a WorkbookModel,
    key: &str,
) -> Result<(MapRef, &'a Sheet), String> {
    let index = keys
        .iter()
        .position(|candidate| candidate == key)
        .ok_or_else(|| format!("sheet {key} is not active"))?;
    let sheet_map = sheets
        .get(txn, key)
        .and_then(|value| value.cast::<MapRef>().ok())
        .ok_or_else(|| format!("missing sheet {key}"))?;
    let sheet_model = model
        .sheets
        .get(index)
        .ok_or_else(|| format!("sheet {key} is missing from the projection"))?;
    Ok((sheet_map, sheet_model))
}

fn sync_sheet(sheet_map: &MapRef, txn: &mut TransactionMut<'_>, sheet: &Sheet) {
    let col_widths: MapRef = sheet_map.get_or_init(txn, COL_WIDTHS);
    let contents: MapRef = sheet_map.get_or_init(txn, CONTENTS);
    sheet_map.try_update(txn, MERGES, merges_to_any(&sheet.merges));
    sheet_map.try_update(txn, NAME, sheet.name.as_str());
    let row_heights: MapRef = sheet_map.get_or_init(txn, ROW_HEIGHTS);
    let styles: MapRef = sheet_map.get_or_init(txn, STYLES);
    sync_numbers(&col_widths, txn, &sheet.col_widths);
    sync_contents(&contents, txn, sheet);
    sync_numbers(&row_heights, txn, &sheet.row_heights);
    sync_styles(&styles, txn, sheet);
}

fn sync_contents(map: &MapRef, txn: &mut TransactionMut<'_>, sheet: &Sheet) {
    let desired = sheet
        .iter_cells()
        .filter_map(|(at, cell)| content_to_any(cell).map(|value| (cell_key(at), value)))
        .collect::<BTreeMap<_, _>>();
    sync_map(map, txn, desired);
}

fn sync_styles(map: &MapRef, txn: &mut TransactionMut<'_>, sheet: &Sheet) {
    let desired = sheet
        .iter_cells()
        .filter_map(|(at, cell)| cell.style.map(|style| (cell_key(at), Any::from(style))))
        .collect::<BTreeMap<_, _>>();
    sync_map(map, txn, desired);
}

fn sync_map(map: &MapRef, txn: &mut TransactionMut<'_>, desired: BTreeMap<String, Any>) {
    let mut stale = map
        .keys(txn)
        .filter(|key| !desired.contains_key(*key))
        .map(str::to_string)
        .collect::<Vec<_>>();
    stale.sort();
    for key in stale {
        map.remove(txn, &key);
    }
    for (key, value) in desired {
        map.try_update(txn, key, value);
    }
}

fn sync_authored_cell(
    sheet_map: &MapRef,
    txn: &mut TransactionMut<'_>,
    sheet: &Sheet,
    at: CellRef,
) {
    let contents: MapRef = sheet_map.get_or_init(txn, CONTENTS);
    let styles: MapRef = sheet_map.get_or_init(txn, STYLES);
    let key = cell_key(at);
    match sheet.cell(at) {
        Some(cell) => {
            sync_optional(&contents, txn, &key, content_to_any(cell));
            sync_optional(&styles, txn, &key, cell.style.map(Any::from));
        }
        None => {
            contents.remove(txn, &key);
            styles.remove(txn, &key);
        }
    }
}

fn sync_optional(map: &MapRef, txn: &mut TransactionMut<'_>, key: &str, value: Option<Any>) {
    if let Some(value) = value {
        map.try_update(txn, key, value);
    } else {
        map.remove(txn, key);
    }
}

fn sync_numbers(map: &MapRef, txn: &mut TransactionMut<'_>, values: &BTreeMap<u32, f64>) {
    let retained = values.keys().map(u32::to_string).collect::<HashSet<_>>();
    let mut stale = map
        .keys(txn)
        .filter(|key| !retained.contains(*key))
        .map(str::to_string)
        .collect::<Vec<_>>();
    stale.sort();
    for key in stale {
        map.remove(txn, &key);
    }
    for (&index, &value) in values {
        map.try_update(txn, index.to_string(), value);
    }
}

fn sync_number(map: &MapRef, txn: &mut TransactionMut<'_>, index: u32, value: Option<f64>) {
    let key = index.to_string();
    match value {
        Some(value) => {
            map.try_update(txn, key, value);
        }
        None => {
            map.remove(txn, &key);
        }
    }
}

fn materialize_sheet<T: ReadTxn>(sheet_map: &MapRef, txn: &T) -> Result<Sheet, String> {
    let name = sheet_map
        .get(txn, NAME)
        .and_then(|value| value.cast::<String>().ok())
        .ok_or_else(|| "sheet is missing its name".to_string())?;
    let mut sheet = Sheet::new(name);
    let mut cells = BTreeMap::<(u32, u32), Cell>::new();
    let contents = nested_map(sheet_map, txn, CONTENTS)?;
    for (key, value) in contents.iter(txn) {
        let at = parse_cell_key(key)?;
        let Out::Any(value) = value else {
            return Err(format!("cell content {key} is not an atomic value"));
        };
        cells.insert((at.row, at.col), content_from_any(&value)?);
    }
    let styles = nested_map(sheet_map, txn, STYLES)?;
    for (key, value) in styles.iter(txn) {
        let at = parse_cell_key(key)?;
        let Out::Any(value) = value else {
            return Err(format!("cell style {key} is not an atomic value"));
        };
        cells.entry((at.row, at.col)).or_default().style = Some(any_u32(&value, "cell style")?);
    }
    for ((row, col), cell) in cells {
        sheet.set_cell(CellRef::new(row, col), cell);
    }
    sheet.col_widths = materialize_numbers(
        &nested_map(sheet_map, txn, COL_WIDTHS)?,
        txn,
        MAX_COLS,
        "column width",
    )?;
    sheet.row_heights = materialize_numbers(
        &nested_map(sheet_map, txn, ROW_HEIGHTS)?,
        txn,
        MAX_ROWS,
        "row height",
    )?;
    sheet.merges = match sheet_map.get(txn, MERGES) {
        Some(Out::Any(value)) => merges_from_any(&value)?,
        _ => return Err("sheet is missing merges".to_string()),
    };
    Ok(sheet)
}

fn nested_map<T: ReadTxn>(parent: &MapRef, txn: &T, key: &str) -> Result<MapRef, String> {
    parent
        .get(txn, key)
        .and_then(|value| value.cast::<MapRef>().ok())
        .ok_or_else(|| format!("sheet is missing {key}"))
}

fn sheet_shared_types<T: ReadTxn>(sheet_map: &MapRef, txn: &T) -> Result<SheetSharedTypes, String> {
    Ok(SheetSharedTypes {
        sheet: sheet_map.as_ref().id(),
        col_widths: nested_map(sheet_map, txn, COL_WIDTHS)?.as_ref().id(),
        contents: nested_map(sheet_map, txn, CONTENTS)?.as_ref().id(),
        row_heights: nested_map(sheet_map, txn, ROW_HEIGHTS)?.as_ref().id(),
        styles: nested_map(sheet_map, txn, STYLES)?.as_ref().id(),
    })
}

fn materialize_numbers<T: ReadTxn>(
    map: &MapRef,
    txn: &T,
    limit: u32,
    label: &str,
) -> Result<BTreeMap<u32, f64>, String> {
    let mut values = BTreeMap::new();
    for (key, value) in map.iter(txn) {
        let index = key
            .parse::<u32>()
            .map_err(|_| format!("invalid numeric key {key}"))?;
        if key != index.to_string() {
            return Err(format!("noncanonical numeric key {key}"));
        }
        if index >= limit {
            return Err(format!("{label} key {key} is out of bounds"));
        }
        let value = value
            .cast::<f64>()
            .map_err(|_| format!("invalid numeric value at {key}"))?;
        if !value.is_finite() {
            return Err(format!("nonfinite {label} at {key}"));
        }
        values.insert(index, value);
    }
    Ok(values)
}

fn sheet_keys<T: ReadTxn>(order: &ArrayRef, txn: &T) -> Result<Vec<String>, String> {
    order
        .iter(txn)
        .map(|value| {
            value
                .cast::<String>()
                .map_err(|_| "sheet order contains a non-string key".to_string())
        })
        .collect()
}

fn structure_generation<T: ReadTxn>(meta: &MapRef, txn: &T) -> Result<i64, String> {
    let generation = meta
        .get(txn, STRUCTURE_GENERATION)
        .and_then(|value| value.cast::<i64>().ok())
        .ok_or_else(|| "missing structure generation".to_string())?;
    if generation < 0 {
        return Err("structure generation is negative".to_string());
    }
    Ok(generation)
}

fn require_root_keys<T: ReadTxn>(txn: &T, expected: &[&str]) -> Result<(), String> {
    let actual = txn
        .root_refs()
        .map(|(key, _)| key.to_string())
        .collect::<BTreeSet<_>>();
    let expected = expected
        .iter()
        .map(|key| (*key).to_string())
        .collect::<BTreeSet<_>>();
    if actual == expected {
        Ok(())
    } else {
        Err("collaborative document roots do not match the schema".to_string())
    }
}

fn require_map_keys<T: ReadTxn>(
    map: &MapRef,
    txn: &T,
    expected: &[&str],
    label: &str,
) -> Result<(), String> {
    let actual = map.keys(txn).map(str::to_string).collect::<BTreeSet<_>>();
    let expected = expected
        .iter()
        .map(|key| (*key).to_string())
        .collect::<BTreeSet<_>>();
    if actual == expected {
        Ok(())
    } else {
        Err(format!("{label} keys do not match the schema"))
    }
}

fn cell_key(at: CellRef) -> String {
    format!("{}:{}", at.row, at.col)
}

fn parse_cell_key(key: &str) -> Result<CellRef, String> {
    let (row, col) = key
        .split_once(':')
        .ok_or_else(|| format!("invalid cell key {key}"))?;
    let row = row
        .parse::<u32>()
        .map_err(|_| format!("invalid cell row in {key}"))?;
    let col = col
        .parse::<u32>()
        .map_err(|_| format!("invalid cell column in {key}"))?;
    let at = CellRef::new(row, col);
    if key != cell_key(at) {
        return Err(format!("noncanonical cell key {key}"));
    }
    if row >= MAX_ROWS || col >= MAX_COLS {
        return Err(format!("cell key {key} is out of bounds"));
    }
    Ok(at)
}

fn content_to_any(cell: &Cell) -> Option<Any> {
    if let Some(formula) = &cell.formula {
        Some(any_array(vec![
            Any::BigInt(1),
            Any::from(formula.as_str()),
            value_to_any(&cell.value),
        ]))
    } else if !matches!(&cell.value, CellValue::Empty) {
        Some(any_array(vec![Any::BigInt(0), value_to_any(&cell.value)]))
    } else {
        None
    }
}

fn content_from_any(value: &Any) -> Result<Cell, String> {
    let values = any_values(value, "cell content")?;
    let kind = values
        .first()
        .ok_or_else(|| "cell content is empty".to_string())?;
    match any_i64(kind, "cell content kind")? {
        0 if values.len() == 2 => {
            let value = value_from_any(&values[1])?;
            if matches!(&value, CellValue::Empty) {
                return Err("empty literal cell content must be omitted".to_string());
            }
            Ok(Cell {
                value,
                ..Cell::default()
            })
        }
        1 if values.len() == 3 => {
            let Any::String(formula) = &values[1] else {
                return Err("formula cell content is missing formula text".to_string());
            };
            Ok(Cell {
                value: value_from_any(&values[2])?,
                formula: Some(formula.to_string()),
                ..Cell::default()
            })
        }
        0 | 1 => Err("cell content has the wrong payload length".to_string()),
        kind => Err(format!("unsupported cell content kind {kind}")),
    }
}

fn value_to_any(value: &CellValue) -> Any {
    match value {
        CellValue::Empty => any_array(vec![Any::BigInt(0)]),
        CellValue::Number { value } => any_array(vec![Any::BigInt(1), Any::Number(*value)]),
        CellValue::Text { value } => any_array(vec![Any::BigInt(2), Any::from(value.as_str())]),
        CellValue::Bool { value } => any_array(vec![Any::BigInt(3), Any::Bool(*value)]),
        CellValue::Error { value } => any_array(vec![Any::BigInt(4), Any::from(value.as_str())]),
    }
}

fn value_from_any(value: &Any) -> Result<CellValue, String> {
    let values = any_values(value, "cell value")?;
    let kind = values
        .first()
        .ok_or_else(|| "cell value is empty".to_string())?;
    match any_i64(kind, "cell value kind")? {
        0 if values.len() == 1 => Ok(CellValue::Empty),
        1 if values.len() == 2 => match &values[1] {
            Any::Number(value) if value.is_finite() => Ok(CellValue::Number { value: *value }),
            _ => Err("numeric cell has a non-number value".to_string()),
        },
        2 if values.len() == 2 => match &values[1] {
            Any::String(value) => Ok(CellValue::Text {
                value: value.to_string(),
            }),
            _ => Err("text cell has a non-string value".to_string()),
        },
        3 if values.len() == 2 => match &values[1] {
            Any::Bool(value) => Ok(CellValue::Bool { value: *value }),
            _ => Err("boolean cell has a non-boolean value".to_string()),
        },
        4 if values.len() == 2 => match &values[1] {
            Any::String(value) => Ok(CellValue::Error {
                value: error_from_str(value)?,
            }),
            _ => Err("error cell has a non-string value".to_string()),
        },
        0..=4 => Err("cell value has the wrong payload length".to_string()),
        kind => Err(format!("unsupported cell value kind {kind}")),
    }
}

fn any_array(values: Vec<Any>) -> Any {
    Any::Array(Arc::from(values))
}

fn any_values<'a>(value: &'a Any, label: &str) -> Result<&'a [Any], String> {
    match value {
        Any::Array(values) => Ok(values),
        _ => Err(format!("{label} is not an array")),
    }
}

fn any_i64(value: &Any, label: &str) -> Result<i64, String> {
    match value {
        Any::BigInt(value) => Ok(*value),
        _ => Err(format!("{label} is not an integer")),
    }
}

fn error_from_str(value: &str) -> Result<ErrorValue, String> {
    match value {
        "#DIV/0!" => Ok(ErrorValue::Div0),
        "#N/A" => Ok(ErrorValue::NA),
        "#NAME?" => Ok(ErrorValue::Name),
        "#NULL!" => Ok(ErrorValue::Null),
        "#NUM!" => Ok(ErrorValue::Num),
        "#REF!" => Ok(ErrorValue::Ref),
        "#VALUE!" => Ok(ErrorValue::Value),
        "#SPILL!" => Ok(ErrorValue::Spill),
        _ => Err(format!("unsupported cell error {value}")),
    }
}

fn merges_to_any(merges: &[CellRange]) -> Any {
    Any::Array(Arc::from(
        merges
            .iter()
            .map(|range| {
                any_array(vec![
                    Any::from(range.start.row),
                    Any::from(range.start.col),
                    Any::Bool(range.start.abs_row),
                    Any::Bool(range.start.abs_col),
                    Any::from(range.end.row),
                    Any::from(range.end.col),
                    Any::Bool(range.end.abs_row),
                    Any::Bool(range.end.abs_col),
                ])
            })
            .collect::<Vec<_>>(),
    ))
}

fn merges_from_any(value: &Any) -> Result<Vec<CellRange>, String> {
    let Any::Array(merges) = value else {
        return Err("sheet merges are not an array".to_string());
    };
    merges
        .iter()
        .map(|merge| {
            let Any::Array(values) = merge else {
                return Err("merge entry is not an array".to_string());
            };
            if values.len() != 8 {
                return Err("merge entry must contain eight values".to_string());
            }
            Ok(CellRange {
                start: CellRef {
                    row: any_u32(&values[0], "merge start row")?,
                    col: any_u32(&values[1], "merge start column")?,
                    abs_row: any_bool(&values[2], "merge start absolute row")?,
                    abs_col: any_bool(&values[3], "merge start absolute column")?,
                },
                end: CellRef {
                    row: any_u32(&values[4], "merge end row")?,
                    col: any_u32(&values[5], "merge end column")?,
                    abs_row: any_bool(&values[6], "merge end absolute row")?,
                    abs_col: any_bool(&values[7], "merge end absolute column")?,
                },
            })
        })
        .collect()
}

fn any_u32(value: &Any, label: &str) -> Result<u32, String> {
    match value {
        Any::Number(value)
            if value.is_finite()
                && *value >= 0.0
                && *value <= u32::MAX as f64
                && value.fract() == 0.0 =>
        {
            Ok(*value as u32)
        }
        Any::BigInt(value) if *value >= 0 && *value <= i64::from(u32::MAX) => Ok(*value as u32),
        _ => Err(format!("{label} is not a u32")),
    }
}

fn any_bool(value: &Any, label: &str) -> Result<bool, String> {
    match value {
        Any::Bool(value) => Ok(*value),
        _ => Err(format!("{label} is not a boolean")),
    }
}

fn fingerprint_model(model: &WorkbookModel) -> Result<(String, u64), String> {
    let mut hasher = Sha256::new();
    hasher.update(b"betteroffice-xlsx-yrs-v2");
    let base = serde_json::to_vec(&(model.date_system, &model.shared_strings, &model.styles))
        .map_err(|error| format!("cannot fingerprint workbook base: {error}"))?;
    hash_bytes(&mut hasher, &base);
    hash_u64(&mut hasher, model.sheets.len() as u64);
    for sheet in &model.sheets {
        hash_bytes(&mut hasher, sheet.name.as_bytes());
        hash_u64(&mut hasher, sheet.iter_cells().count() as u64);
        for (at, cell) in sheet.iter_cells() {
            hash_u32(&mut hasher, at.row);
            hash_u32(&mut hasher, at.col);
            hash_cell_value(&mut hasher, &cell.value);
            match &cell.formula {
                Some(formula) => {
                    hasher.update([1]);
                    hash_bytes(&mut hasher, formula.as_bytes());
                }
                None => hasher.update([0]),
            }
            match cell.style {
                Some(style) => {
                    hasher.update([1]);
                    hash_u32(&mut hasher, style);
                }
                None => hasher.update([0]),
            }
        }
        hash_u64(&mut hasher, sheet.merges.len() as u64);
        for range in &sheet.merges {
            hash_cell_ref(&mut hasher, range.start);
            hash_cell_ref(&mut hasher, range.end);
        }
        hash_u64(&mut hasher, sheet.col_widths.len() as u64);
        for (&column, &width) in &sheet.col_widths {
            hash_u32(&mut hasher, column);
            hash_u64(&mut hasher, width.to_bits());
        }
        hash_u64(&mut hasher, sheet.row_heights.len() as u64);
        for (&row, &height) in &sheet.row_heights {
            hash_u32(&mut hasher, row);
            hash_u64(&mut hasher, height.to_bits());
        }
    }
    let digest = hasher.finalize();
    let fingerprint = format!("{digest:x}");
    let mut client_bytes = [0_u8; 8];
    client_bytes.copy_from_slice(&digest[..8]);
    let mut bootstrap_client_id = u64::from_be_bytes(client_bytes) & MAX_SAFE_CLIENT_ID;
    if bootstrap_client_id == 0 {
        bootstrap_client_id = 1;
    }
    Ok((fingerprint, bootstrap_client_id))
}

fn hash_cell_value(hasher: &mut Sha256, value: &CellValue) {
    match value {
        CellValue::Empty => hasher.update([0]),
        CellValue::Number { value } => {
            hasher.update([1]);
            hash_u64(hasher, value.to_bits());
        }
        CellValue::Text { value } => {
            hasher.update([2]);
            hash_bytes(hasher, value.as_bytes());
        }
        CellValue::Bool { value } => hasher.update([3, u8::from(*value)]),
        CellValue::Error { value } => {
            hasher.update([4]);
            hash_bytes(hasher, value.as_str().as_bytes());
        }
    }
}

fn hash_cell_ref(hasher: &mut Sha256, cell: CellRef) {
    hash_u32(hasher, cell.row);
    hash_u32(hasher, cell.col);
    hasher.update([u8::from(cell.abs_row), u8::from(cell.abs_col)]);
}

fn hash_bytes(hasher: &mut Sha256, bytes: &[u8]) {
    hash_u64(hasher, bytes.len() as u64);
    hasher.update(bytes);
}

fn hash_u32(hasher: &mut Sha256, value: u32) {
    hasher.update(value.to_le_bytes());
}

fn hash_u64(hasher: &mut Sha256, value: u64) {
    hasher.update(value.to_le_bytes());
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::Xf;

    fn rich_model() -> WorkbookModel {
        let mut first = Sheet::new("Data");
        first.set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::Number { value: 42.0 },
                formula: Some("40+2".into()),
                style: Some(0),
            },
        );
        first.set_cell(
            CellRef::new(1, 0),
            Cell {
                value: CellValue::Text {
                    value: "hello".into(),
                },
                ..Cell::default()
            },
        );
        first.col_widths.insert(1, 24.5);
        first.row_heights.insert(2, 30.0);
        first
            .merges
            .push(CellRange::new(CellRef::new(3, 0), CellRef::new(4, 2)));
        let mut model = WorkbookModel {
            date_system: DateSystem::V1904,
            shared_strings: vec!["hello".into()],
            ..WorkbookModel::default()
        };
        model.styles.cell_xfs.push(Xf::default());
        model.sheets.push(first);
        model.sheets.push(Sheet::new("Second"));
        model
    }

    #[test]
    fn deterministic_bootstrap_round_trips_formula_fallbacks() {
        let model = rich_model();
        let left = WorkbookAuthority::from_model_with_client_id(&model, 11).unwrap();
        let right = WorkbookAuthority::from_model_with_client_id(&model, 12).unwrap();
        assert_eq!(left.materialize().unwrap(), model);
        assert_eq!(right.materialize().unwrap(), model);
        assert_eq!(
            left.encode_state_vector_v1(),
            right.encode_state_vector_v1()
        );
        assert_eq!(
            left.encode_state_as_update_v1(),
            right.encode_state_as_update_v1()
        );
    }

    #[test]
    fn formula_content_is_one_atomic_payload() {
        let formula = Cell {
            value: CellValue::Error {
                value: ErrorValue::Ref,
            },
            formula: Some("Missing!A1".into()),
            style: None,
        };
        assert_eq!(
            content_from_any(&content_to_any(&formula).unwrap()).unwrap(),
            formula
        );
    }

    #[test]
    fn strict_decoders_reject_trailing_and_impossible_vectors() {
        assert!(decode_state_vector_v1(&[1]).is_err());
        assert!(decode_state_vector_v1(&[0, 0]).is_err());
        assert!(decode_update_v1(&[0, 0, 0]).is_err());
    }

    #[test]
    fn state_vector_entry_count_is_bounded_before_decode() {
        let error = decode_state_vector_v1(&[0x81, 0x80, 0x04]).unwrap_err();
        assert!(error.contains("65536-entry limit"), "{error}");
    }

    #[test]
    fn bootstrap_client_id_conflicts_are_rejected_locally() {
        let model = rich_model();
        let base = WorkbookBase::from_model(&model).unwrap();
        assert!(matches!(
            WorkbookAuthority::from_model_with_client_id(&model, base.bootstrap_client_id),
            Err(AuthorityError::ClientIdConflict(_))
        ));
    }

    #[test]
    fn shared_map_replacement_changes_the_frozen_structure() {
        let model = rich_model();
        let source = WorkbookAuthority::from_model_with_client_id(&model, 21).unwrap();
        let target = WorkbookAuthority::from_model_with_client_id(&model, 22).unwrap();
        let target_structure = target.structure().unwrap();
        let target_vector = target.encode_state_vector_v1();

        {
            let mut txn = source.doc.transact_mut_with("test:replace-map");
            let sheets = txn.get_map(SHEETS).unwrap();
            let sheet = sheets
                .get(&txn, "sheet:1")
                .and_then(|value| value.cast::<MapRef>().ok())
                .unwrap();
            sheet.insert(&mut txn, CONTENTS, MapPrelim::default());
        }

        let update = source.encode_diff_v1(&target_vector).unwrap();
        let staged = target.stage_updates_v1(&[&update]).unwrap();
        assert_eq!(staged.model, target.materialize().unwrap());
        assert_ne!(staged.structure, target_structure);
    }

    #[test]
    fn retained_sheet_maps_stay_valid_and_keep_identity_through_undo() {
        let model = rich_model();
        let mut authority = WorkbookAuthority::from_model_with_client_id(&model, 31).unwrap();
        authority
            .apply_ops(&[Op::RemoveSheet { index: 1 }], SyncOrigin::User)
            .unwrap();
        let (_, removed) = authority.strict_materialize().unwrap();
        assert_eq!(removed.sheet_keys, ["sheet:0"]);
        assert_eq!(removed.shared_types.len(), 2);
        let retained = removed.shared_types["sheet:1"].clone();

        authority
            .apply_ops(
                &[Op::AddSheet {
                    index: 1,
                    name: "Second".into(),
                }],
                SyncOrigin::Undo,
            )
            .unwrap();
        let (restored, structure) = authority.strict_materialize().unwrap();
        assert_eq!(restored, model);
        assert_eq!(structure.shared_types["sheet:1"], retained);
    }
}
