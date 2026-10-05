use std::collections::VecDeque;
use std::mem::size_of;
use std::num::NonZeroUsize;

use ooxml_opc::{SourceContainer, SourceContainerBuilder};
use sha2::{Digest, Sha256};
use xlsx_calc::graph::DepGraphBuilder;
use xlsx_model::{CellRange, CellRef, Workbook as WorkbookModel};
use xlsx_parse::{PackageFactsBuilder, PackageFactsEncoder};

use crate::authority::snapshot::{AuthorityHydrator, AuthoritySnapshotEncoder};
use crate::snapshot::header::{SnapshotHeader, SnapshotMode};
use crate::snapshot::model::{ModelSnapshotBuilder, ModelSnapshotEncoder};
use crate::snapshot::preserved::{PreservedSnapshotBuilder, PreservedSnapshotEncoder};
use crate::snapshot::wire::{ChunkKind, Reader, Writer, frame, unframe};
use crate::snapshot::{SnapshotBudget, SnapshotError, SnapshotProgress, SnapshotResult};

use super::{
    Arc, BTreeMap, ChartCache, ChartAnchor, HashMap, Mutex, PackageSlot, PreservedSheetState,
    ProposalSet, SheetId, UndoStack, UpdateObservers, Workbook, WorkbookMode,
};
use crate::{CalculationOptions, CalculationResult, CellAddress};

const MAX_LOGICAL_BYTES: usize = 64 * 1024 * 1024;

fn index(kind: ChunkKind) -> usize {
    kind as usize - 1
}

fn error(message: impl Into<String>) -> SnapshotError {
    SnapshotError::new(message)
}

struct Lineage {
    nonce: String,
    changes: u64,
    epoch: u64,
    active_sheet: SheetId,
    seed: Option<u32>,
    recalculated: bool,
    projection_valid: bool,
    state_vector: Vec<u8>,
}

impl Lineage {
    fn capture(workbook: &Workbook) -> SnapshotResult<Self> {
        let Workbook {
            authority,
            mode: _,
            pending_remote_updates,
            model,
            source_package,
            source_container,
            preserved,
            preserved_undo,
            preserved_redo,
            edited_since_open,
            recalculated_since_open,
            moved_references_since_open,
            active_sheet,
            undo,
            graph: _,
            rand_seed,
            proposals,
            last_calculation: _,
            update_observers,
            opened_anchors: _,
            sheet_info_cache: _,
            model_epoch,
            geometry_cache: _,
            version_nonce,
            committed_changes,
            chart_cache: _,
            source_part_hashes: _,
        } = workbook;
        workbook.require_snapshot_standalone()?;
        if *edited_since_open
            || *moved_references_since_open
            || undo.can_undo()
            || undo.can_redo()
            || !preserved_undo.is_empty()
            || !preserved_redo.is_empty()
            || !pending_remote_updates.is_empty()
            || !proposals.list().is_empty()
        {
            return Err(error("snapshot requires an initial workbook"));
        }
        let mut proposal_counter = proposals.clone();
        if proposal_counter.next_id() != "p1" {
            return Err(error("snapshot requires an initial proposal counter"));
        }
        let observers = update_observers
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if !observers.listeners.is_empty() || observers.next_id != 0 {
            return Err(error("snapshot requires an initial observer registry"));
        }
        if source_package.is_some() && source_container.is_none() {
            return Err(error("snapshot source package has no original container"));
        }
        if matches!(source_package, Some(PackageSlot::Deferred { .. })) {
            return Err(error("snapshot capture requires a present source package"));
        }
        let PreservedSheetState {
            origins,
            shared_string_cells,
            axes,
            created,
        } = preserved;
        let sheets = model.sheets.len();
        if [origins.len(), shared_string_cells.len(), axes.len(), created.len()]
            .iter()
            .any(|count| *count != sheets)
            || created.iter().any(|created| *created)
            || axes.iter().flatten().any(|axes| {
                let xlsx_parse::SheetAxes { rows, cols } = axes;
                !rows.is_identity()
                    || !cols.is_identity()
                    || axes != &xlsx_parse::SheetAxes::default()
            })
        {
            return Err(error("snapshot preservation state is not initial"));
        }
        model.styles.snapshot_field_counts();
        Ok(Self {
            nonce: version_nonce.clone(),
            changes: *committed_changes,
            epoch: *model_epoch,
            active_sheet: *active_sheet,
            seed: *rand_seed,
            recalculated: *recalculated_since_open,
            projection_valid: authority.snapshot_projection_valid()?,
            state_vector: authority.encode_state_vector_v1(),
        })
    }

    fn matches(&self, workbook: &Workbook) -> SnapshotResult<bool> {
        workbook.require_snapshot_standalone()?;
        if self.nonce != workbook.version_nonce
            || self.changes != workbook.committed_changes
            || self.epoch != workbook.model_epoch
            || self.active_sheet != workbook.active_sheet
            || self.seed != workbook.rand_seed
            || self.recalculated != workbook.recalculated_since_open
            || workbook.edited_since_open
            || workbook.moved_references_since_open
            || !workbook.pending_remote_updates.is_empty()
            || workbook.undo.can_undo()
            || workbook.undo.can_redo()
            || !workbook.preserved_undo.is_empty()
            || !workbook.preserved_redo.is_empty()
            || !workbook.proposals.list().is_empty()
        {
            return Ok(false);
        }
        let observers = workbook
            .update_observers
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if !observers.listeners.is_empty() || observers.next_id != 0 {
            return Ok(false);
        }
        let mut proposals = workbook.proposals.clone();
        Ok(proposals.next_id() == "p1"
            && self.projection_valid == workbook.authority.snapshot_projection_valid()?
            && self.state_vector == workbook.authority.encode_state_vector_v1())
    }
}

#[doc(hidden)]
pub struct WorkbookSnapshotEncoder {
    lineage: Lineage,
    snapshot_id: u64,
    budget: SnapshotBudget,
    part_budget: SnapshotBudget,
    header: Option<Vec<u8>>,
    authority: AuthoritySnapshotEncoder,
    model: ModelSnapshotEncoder,
    preserved: PreservedSnapshotEncoder,
    facts: PackageFactsEncoder,
    source: Option<SourceContainer>,
    source_offset: usize,
    source_ordinal: u64,
    facts_ordinal: u64,
    calculation_list: usize,
    calculation_index: usize,
    header_ordinal: u64,
    stage: u8,
    pending: Option<Vec<u8>>,
    offset: usize,
    ordinals: [u64; 9],
    digest: Sha256,
    split_fallback: Option<String>,
    failed: bool,
}

impl WorkbookSnapshotEncoder {
    pub fn new(
        workbook: &Workbook,
        context: Option<CalculationOptions>,
        budget: SnapshotBudget,
    ) -> SnapshotResult<Self> {
        let lineage = Lineage::capture(workbook)?;
        if budget.max_bytes() < 64 {
            return Err(error("snapshot byte budget is too small for framing"));
        }
        let snapshot_id = u64::from_str_radix(&yrs::uuid_v4().replace('-', "")[..16], 16)
            .map_err(|_| error("cannot allocate snapshot identity"))?;
        let part_bytes = budget
            .max_bytes()
            .saturating_sub(64)
            .clamp(1, MAX_LOGICAL_BYTES - 12);
        let part_budget = SnapshotBudget::new(1, part_bytes)?;
        let record_budget = SnapshotBudget::new(1, MAX_LOGICAL_BYTES)?;
        let authority = AuthoritySnapshotEncoder::new(&workbook.authority, part_budget)?;
        let split_fallback = authority.split_fallback();
        let (base_count, yrs_count) = authority.chunk_counts(part_budget);
        let CalculationResult {
            changed,
            cycle_cells,
            limited_cells,
        } = &workbook.last_calculation;
        let calculation_counts = [changed.len(), cycle_cells.len(), limited_cells.len()];
        let header_count = calculation_counts.iter().try_fold(1u64, |total, count| {
            total
                .checked_add(*count as u64)
                .ok_or_else(|| error("snapshot header count overflows"))
        })?;
        let mut chunk_counts = [header_count, base_count, yrs_count, 0, 0, 0, 0, 0, 1];
        let mut model = ModelSnapshotEncoder::new();
        while let Some(chunk) = model.next(&workbook.model, record_budget)? {
            count_chunk(&mut chunk_counts, &chunk)?;
        }
        let mut preserved = PreservedSnapshotEncoder::new();
        while let Some(chunk) = preserved.next(&workbook.preserved, record_budget)? {
            count_chunk(&mut chunk_counts, &chunk)?;
        }
        if let Some(PackageSlot::Present(package)) = &workbook.source_package {
            let mut facts = PackageFactsEncoder::new();
            while let Some(payload) = facts
                .next(package, part_budget.max_bytes())
                .map_err(|failure| error(failure.to_string()))?
            {
                let ordinal = chunk_counts[index(ChunkKind::Facts)];
                count_chunk(&mut chunk_counts, &frame(ChunkKind::Facts, ordinal, &payload))?;
            }
        }
        if let Some(source) = &workbook.source_container {
            chunk_counts[index(ChunkKind::Source)] = source
                .as_bytes()
                .len()
                .div_ceil(part_budget.max_bytes())
                .max(1) as u64;
        }
        let (client_id, guid, next_sheet_id) = workbook.authority.snapshot_identity();
        let header = SnapshotHeader {
            snapshot_id,
            mode: SnapshotMode::Standalone,
            edited_since_open: workbook.edited_since_open,
            recalculated_since_open: workbook.recalculated_since_open,
            moved_references_since_open: workbook.moved_references_since_open,
            active_sheet: workbook.active_sheet,
            rand_seed: workbook.rand_seed,
            model_epoch: workbook.model_epoch,
            version_nonce: workbook.version_nonce.clone(),
            committed_changes: workbook.committed_changes,
            last_calculation: CalculationResult::default(),
            calculation_context: context,
            client_id,
            guid,
            next_sheet_id,
            state_vector: lineage.state_vector.clone(),
            chunk_counts,
        };
        let mut payload = Writer::new();
        payload.bytes(&header.encode());
        payload.bool(lineage.projection_valid);
        payload.bool(workbook.graph.is_some());
        payload.bool(workbook.source_package.is_some());
        payload.option(workbook.source_container.as_ref(), |writer, source| {
            writer.var_usize(source.as_bytes().len());
        });
        for count in workbook.model.styles.snapshot_field_counts() {
            payload.var_usize(count);
        }
        for count in calculation_counts {
            payload.var_usize(count);
        }
        let header_chunk = frame(ChunkKind::Header, 0, &payload.into_bytes());
        if header_chunk.len() > MAX_LOGICAL_BYTES {
            return Err(error("snapshot header exceeds hydration limit"));
        }
        Ok(Self {
            lineage,
            snapshot_id,
            budget,
            part_budget,
            header: Some(header_chunk),
            authority,
            model: ModelSnapshotEncoder::new(),
            preserved: PreservedSnapshotEncoder::new(),
            facts: PackageFactsEncoder::new(),
            source: workbook.source_container.clone(),
            source_offset: 0,
            source_ordinal: 0,
            facts_ordinal: 0,
            calculation_list: 0,
            calculation_index: 0,
            header_ordinal: 1,
            stage: 0,
            pending: None,
            offset: 0,
            ordinals: [0; 9],
            digest: Sha256::new(),
            split_fallback,
            failed: false,
        })
    }

    pub fn next(&mut self, workbook: &Workbook) -> SnapshotResult<Option<Vec<u8>>> {
        if self.failed {
            return Err(error("snapshot encoder has failed"));
        }
        let result = self.next_inner(workbook);
        self.failed = result.is_err();
        result
    }

    fn next_inner(&mut self, workbook: &Workbook) -> SnapshotResult<Option<Vec<u8>>> {
        if !self.lineage.matches(workbook)? {
            return Err(error("snapshot workbook lineage has changed"));
        }
        if self.pending.is_none() {
            self.pending = self.next_logical(workbook)?;
        }
        let Some(chunk) = &self.pending else {
            return Ok(None);
        };
        let (kind, ordinal, _) = unframe(chunk)?;
        if chunk.len() > MAX_LOGICAL_BYTES {
            return Err(error(format!(
                "snapshot {kind:?} chunk exceeds hydration limit: {}",
                self.split_fallback.as_deref().unwrap_or("oversized_record"),
            )));
        }
        let mut payload = Writer::new();
        payload.var_u64(self.snapshot_id);
        payload.var_u64(ordinal);
        payload.var_usize(chunk.len());
        payload.var_usize(self.offset);
        let prefix = frame(kind, self.ordinals[index(kind)], &payload.into_bytes());
        let available = self
            .budget
            .max_bytes()
            .checked_sub(prefix.len())
            .filter(|available| *available != 0)
            .ok_or_else(|| error("snapshot byte budget is too small for framing"))?;
        let end = self.offset + available.min(chunk.len() - self.offset);
        let mut transport = prefix;
        transport.extend_from_slice(&chunk[self.offset..end]);
        if kind != ChunkKind::End {
            self.digest.update(&chunk[self.offset..end]);
        }
        self.ordinals[index(kind)] += 1;
        if end == chunk.len() {
            self.pending = None;
            self.offset = 0;
        } else {
            self.offset = end;
        }
        Ok(Some(transport))
    }

    fn next_logical(&mut self, workbook: &Workbook) -> SnapshotResult<Option<Vec<u8>>> {
        let record_budget = SnapshotBudget::new(1, MAX_LOGICAL_BYTES)?;
        loop {
            match self.stage {
                0 => {
                    if let Some(header) = self.header.take() {
                        return Ok(Some(header));
                    }
                    let CalculationResult {
                        changed,
                        cycle_cells,
                        limited_cells,
                    } = &workbook.last_calculation;
                    let lists = [changed, cycle_cells, limited_cells];
                    while let Some(list) = lists.get(self.calculation_list) {
                        if let Some(address) = list.get(self.calculation_index) {
                            let mut payload = Writer::new();
                            payload.u8(self.calculation_list as u8);
                            write_address(&mut payload, address);
                            let chunk = frame(
                                ChunkKind::Header,
                                self.header_ordinal,
                                &payload.into_bytes(),
                            );
                            self.header_ordinal += 1;
                            self.calculation_index += 1;
                            return Ok(Some(chunk));
                        }
                        self.calculation_list += 1;
                        self.calculation_index = 0;
                    }
                }
                1 => {
                    if let Some(chunk) = self.authority.next(self.part_budget)? {
                        return Ok(Some(chunk));
                    }
                }
                2 => {
                    if let Some(chunk) = self.model.next(&workbook.model, record_budget)? {
                        return Ok(Some(chunk));
                    }
                }
                3 => {
                    if let Some(chunk) = self.preserved.next(&workbook.preserved, record_budget)? {
                        return Ok(Some(chunk));
                    }
                }
                4 => {
                    if let Some(PackageSlot::Present(package)) = &workbook.source_package
                        && let Some(payload) = self
                            .facts
                            .next(package, self.part_budget.max_bytes())
                            .map_err(|failure| error(failure.to_string()))?
                    {
                        let chunk = frame(ChunkKind::Facts, self.facts_ordinal, &payload);
                        self.facts_ordinal += 1;
                        return Ok(Some(chunk));
                    }
                }
                5 => {
                    if let Some(source) = &self.source
                        && (self.source_offset < source.as_bytes().len()
                            || self.source_ordinal == 0)
                    {
                        let bytes = source.as_bytes();
                        let end = self.source_offset
                            + self.part_budget.max_bytes().min(bytes.len() - self.source_offset);
                        let chunk = frame(
                            ChunkKind::Source,
                            self.source_ordinal,
                            &bytes[self.source_offset..end],
                        );
                        self.source_offset = end;
                        self.source_ordinal += 1;
                        return Ok(Some(chunk));
                    }
                }
                6 => {
                    self.stage = 7;
                    return Ok(Some(frame(ChunkKind::End, 0, &self.digest.clone().finalize())));
                }
                _ => return Ok(None),
            }
            self.stage += 1;
        }
    }

    pub fn split_fallback_reason(&self) -> Option<&str> {
        self.split_fallback.as_deref()
    }
}

fn count_chunk(counts: &mut [u64; 9], chunk: &[u8]) -> SnapshotResult<()> {
    if chunk.len() > MAX_LOGICAL_BYTES {
        return Err(error("snapshot logical chunk exceeds hydration limit"));
    }
    let (kind, _, _) = unframe(chunk)?;
    counts[index(kind)] = counts[index(kind)]
        .checked_add(1)
        .ok_or_else(|| error("snapshot chunk count overflows"))?;
    Ok(())
}

struct Fragment {
    kind: ChunkKind,
    ordinal: u64,
    length: usize,
    bytes: Vec<u8>,
}

fn write_address(writer: &mut Writer, address: &CellAddress) {
    let CellAddress { sheet, cell } = address;
    let SheetId(sheet) = sheet;
    let CellRef {
        row,
        col,
        abs_row,
        abs_col,
    } = cell;
    writer.var_u32(*sheet);
    writer.var_u32(*row);
    writer.var_u32(*col);
    writer.bool(*abs_row);
    writer.bool(*abs_col);
}

fn read_address(reader: &mut Reader<'_>) -> SnapshotResult<CellAddress> {
    Ok(CellAddress {
        sheet: SheetId(reader.var_u32()?),
        cell: CellRef {
            row: reader.var_u32()?,
            col: reader.var_u32()?,
            abs_row: reader.bool()?,
            abs_col: reader.bool()?,
        },
    })
}

struct HeaderState {
    header: SnapshotHeader,
    projection_valid: bool,
    graph_present: bool,
    package_present: bool,
    source_length: Option<usize>,
    style_counts: [usize; 7],
    calculation_counts: [usize; 3],
    calculation_list: usize,
    calculation_index: usize,
}

#[derive(Clone, Copy, Default)]
struct StreamAdmission {
    remaining: usize,
    length: usize,
    varint: u64,
    shift: u32,
}

impl StreamAdmission {
    fn push(&mut self, mut payload: &[u8], max_bytes: usize) -> SnapshotResult<()> {
        if self.remaining != 0 && self.length > max_bytes.min(MAX_LOGICAL_BYTES) {
            return Err(error("snapshot record exceeds advance byte budget"));
        }
        while !payload.is_empty() {
            if self.remaining != 0 {
                let count = self.remaining.min(payload.len());
                self.remaining -= count;
                payload = &payload[count..];
            } else {
                let byte = payload[0];
                payload = &payload[1..];
                let bits = u64::from(byte & 0x7f);
                if self.shift == 63 && bits > 1 {
                    return Err(error("snapshot record length overflows"));
                }
                self.varint |= bits << self.shift;
                if byte & 0x80 == 0 {
                    let length = usize::try_from(self.varint)
                        .map_err(|_| error("snapshot record length overflows"))?;
                    if length == 0 {
                        return Err(error("snapshot record is empty"));
                    }
                    if length > max_bytes.min(MAX_LOGICAL_BYTES) {
                        return Err(error("snapshot record exceeds advance byte budget"));
                    }
                    self.remaining = length;
                    self.length = length;
                    self.varint = 0;
                    self.shift = 0;
                } else {
                    self.shift += 7;
                    if self.shift > 63 {
                        return Err(error("snapshot record length overflows"));
                    }
                }
            }
        }
        Ok(())
    }
}

#[derive(Clone, Copy, Default)]
struct GraphCursor {
    phase: u8,
    index: usize,
    after: Option<CellRef>,
}

impl GraphCursor {
    fn next_cost(&mut self, model: &WorkbookModel) -> Option<usize> {
        loop {
            match self.phase {
                0 => {
                    if let Some(sheet) = model.sheets.get(self.index) {
                        self.index += 1;
                        return Some(sheet.name.len().saturating_add(32));
                    }
                }
                1 => {
                    if let Some(name) = model.defined_names.get(self.index) {
                        self.index += 1;
                        return Some(name.name.len().saturating_add(name.formula.len()));
                    }
                }
                2 => {
                    if let Some(table) = model.tables.get(self.index) {
                        self.index += 1;
                        return Some(table.columns.iter().fold(table.name.len(), |bytes, column| {
                            bytes.saturating_add(column.len())
                        }));
                    }
                }
                3 => {
                    let Some(sheet) = model.sheets.get(self.index) else {
                        return None;
                    };
                    let row = self.after.map_or(0, |at| at.row);
                    let next_col = self.after.and_then(|at| at.col.checked_add(1));
                    let current_row = next_col.and_then(|col| {
                        sheet
                            .cells_in_range(CellRange::new(
                                CellRef::new(row, col),
                                CellRef::new(row, u32::MAX),
                            ))
                            .next()
                    });
                    let start_row = if self.after.is_none() {
                        Some(0)
                    } else {
                        row.checked_add(1)
                    };
                    let next = current_row.or_else(|| {
                        start_row.and_then(|row| {
                            sheet
                                .cells_in_range(CellRange::new(
                                    CellRef::new(row, 0),
                                    CellRef::new(u32::MAX, u32::MAX),
                                ))
                                .next()
                        })
                    });
                    if let Some((at, cell)) = next {
                        self.after = Some(at);
                        return Some(cell.formula.as_ref().map_or(1, String::len));
                    }
                    self.index += 1;
                    self.after = None;
                    return Some(1);
                }
                _ => return None,
            }
            self.phase += 1;
            self.index = 0;
        }
    }
}

pub struct WorkbookSnapshotBuilder {
    header: Option<HeaderState>,
    snapshot_id: Option<u64>,
    ordinals: [u64; 9],
    received: [u64; 9],
    fragment: Option<Fragment>,
    queue: VecDeque<Vec<u8>>,
    digest: Sha256,
    ended: bool,
    failed: bool,
    authority: Option<AuthorityHydrator>,
    model: Option<ModelSnapshotBuilder>,
    preserved: Option<PreservedSnapshotBuilder>,
    facts: Option<PackageFactsBuilder>,
    source: Option<SourceContainerBuilder>,
    restored: Option<Workbook>,
    graph: Option<DepGraphBuilder>,
    anchor_sheet: usize,
    anchor_chart: usize,
    anchors_done: bool,
    ready: Option<HydratedWorkbook>,
    calculation_context: Option<CalculationOptions>,
    base_admission: StreamAdmission,
    facts_admission: StreamAdmission,
    graph_cursor: GraphCursor,
}

impl WorkbookSnapshotBuilder {
    pub fn new() -> Self {
        Self {
            header: None,
            snapshot_id: None,
            ordinals: [0; 9],
            received: [0; 9],
            fragment: None,
            queue: VecDeque::new(),
            digest: Sha256::new(),
            ended: false,
            failed: false,
            authority: None,
            model: None,
            preserved: None,
            facts: None,
            source: None,
            restored: None,
            graph: None,
            anchor_sheet: 0,
            anchor_chart: 0,
            anchors_done: false,
            ready: None,
            calculation_context: None,
            base_admission: StreamAdmission::default(),
            facts_admission: StreamAdmission::default(),
            graph_cursor: GraphCursor::default(),
        }
    }

    pub fn push(&mut self, chunk: &[u8]) -> SnapshotResult<SnapshotProgress> {
        if self.failed {
            return Err(error("snapshot builder has failed"));
        }
        let result = self.push_inner(chunk);
        self.failed = result.is_err();
        result
    }

    fn push_inner(&mut self, chunk: &[u8]) -> SnapshotResult<SnapshotProgress> {
        if self.ended || self.ready.is_some() {
            return Err(error("snapshot has extra chunks"));
        }
        let (kind, ordinal, payload) = unframe(chunk)?;
        if ordinal != self.ordinals[index(kind)] {
            return Err(error("snapshot transport ordinal differs"));
        }
        let mut reader = Reader::new(payload);
        let snapshot_id = reader.var_u64()?;
        let logical_ordinal = reader.var_u64()?;
        let length = reader.var_usize()?;
        let offset = reader.var_usize()?;
        let bytes = reader.rest();
        if self.snapshot_id.is_some_and(|expected| expected != snapshot_id) {
            return Err(error("snapshot lineage differs"));
        }
        if bytes.is_empty() || length == 0 || length > MAX_LOGICAL_BYTES {
            return Err(error("invalid snapshot fragment length"));
        }
        if self.fragment.is_none() {
            let expected_kind = match &self.header {
                None if self.received[index(ChunkKind::Header)] == 0 => index(ChunkKind::Header),
                None => {
                    let last = self.received.iter().rposition(|count| *count != 0).unwrap();
                    if index(kind) < last {
                        return Err(error("snapshot chunks are reordered"));
                    }
                    index(kind)
                }
                Some(state) => self
                    .received
                    .iter()
                    .zip(state.header.chunk_counts)
                    .position(|(received, expected)| *received < expected)
                    .ok_or_else(|| error("snapshot has extra chunks"))?,
            };
            if index(kind) != expected_kind || offset != 0 {
                return Err(error("snapshot chunks are missing or reordered"));
            }
            let mut storage = Vec::new();
            storage
                .try_reserve_exact(length)
                .map_err(|_| error("cannot allocate snapshot fragment"))?;
            self.fragment = Some(Fragment {
                kind,
                ordinal: logical_ordinal,
                length,
                bytes: storage,
            });
        }
        let fragment = self.fragment.as_mut().expect("snapshot fragment exists");
        if fragment.kind != kind
            || fragment.ordinal != logical_ordinal
            || fragment.length != length
            || fragment.bytes.len() != offset
            || bytes.len() > length - offset
        {
            return Err(error("snapshot fragments are missing or reordered"));
        }
        fragment
            .bytes
            .try_reserve(bytes.len())
            .map_err(|_| error("cannot allocate snapshot fragment"))?;
        fragment.bytes.extend_from_slice(bytes);
        if kind != ChunkKind::End {
            self.digest.update(bytes);
        }
        self.snapshot_id = Some(snapshot_id);
        self.ordinals[index(kind)] = ordinal
            .checked_add(1)
            .ok_or_else(|| error("snapshot ordinal overflows"))?;
        if fragment.bytes.len() != length {
            return Ok(SnapshotProgress::pending());
        }
        let fragment = self.fragment.take().expect("complete snapshot fragment");
        let (inner_kind, inner_ordinal, payload) = unframe(&fragment.bytes)?;
        if let Some(state) = &self.header {
            let expected = self
                .received
                .iter()
                .zip(state.header.chunk_counts)
                .position(|(received, expected)| *received < expected)
                .ok_or_else(|| error("snapshot has extra chunks"))?;
            if index(kind) != expected {
                return Err(error("snapshot chunks are missing or reordered"));
            }
        }
        let expected_ordinal = if kind == ChunkKind::Cells {
            self.received[index(ChunkKind::Model)] + self.received[index(ChunkKind::Cells)]
        } else {
            self.received[index(kind)]
        };
        if inner_kind != kind
            || inner_ordinal != logical_ordinal
            || inner_ordinal != expected_ordinal
        {
            return Err(error("snapshot logical ordinal differs"));
        }
        match kind {
            ChunkKind::End => {
                let actual = self.digest.clone().finalize();
                if payload != &actual[..] {
                    return Err(error("snapshot content digest differs"));
                }
                self.ended = true;
            }
            _ => self.queue.push_back(fragment.bytes),
        }
        self.received[index(kind)] += 1;
        Ok(SnapshotProgress::pending())
    }

    fn accept_header(&mut self, payload: &[u8], snapshot_id: u64) -> SnapshotResult<()> {
        let mut reader = Reader::new(payload);
        let header = SnapshotHeader::decode(reader.bytes()?)?;
        if header.snapshot_id != snapshot_id
            || header.edited_since_open
            || header.moved_references_since_open
            || header.next_sheet_id != 0
            || header.chunk_count(ChunkKind::Header) == 0
            || header.chunk_count(ChunkKind::End) != 1
            || header.chunk_count(ChunkKind::Model) == 0
            || header.chunk_count(ChunkKind::Preserved) == 0
        {
            return Err(error("snapshot header is invalid"));
        }
        let projection_valid = reader.bool()?;
        let graph_present = reader.bool()?;
        let package_present = reader.bool()?;
        let source_length = reader.option(Reader::var_usize)?;
        let mut style_counts = [0; 7];
        for count in &mut style_counts {
            *count = reader.var_usize()?;
        }
        let mut calculation_counts = [0; 3];
        for count in &mut calculation_counts {
            *count = reader.var_usize()?;
        }
        reader.finish()?;
        let header_count = calculation_counts.iter().try_fold(1u64, |total, count| {
            total
                .checked_add(*count as u64)
                .ok_or_else(|| error("snapshot header count overflows"))
        })?;
        if header_count != header.chunk_count(ChunkKind::Header)
            || !header.last_calculation.changed.is_empty()
            || !header.last_calculation.cycle_cells.is_empty()
            || !header.last_calculation.limited_cells.is_empty()
        {
            return Err(error("snapshot calculation counts differ"));
        }
        let mut header = header;
        let CalculationResult {
            changed,
            cycle_cells,
            limited_cells,
        } = &mut header.last_calculation;
        for (list, count) in [changed, cycle_cells, limited_cells]
            .into_iter()
            .zip(calculation_counts)
        {
            list.try_reserve_exact(count)
                .map_err(|_| error("cannot allocate snapshot calculation"))?;
        }
        if package_present != (header.chunk_count(ChunkKind::Facts) != 0)
            || source_length.is_some() != (header.chunk_count(ChunkKind::Source) != 0)
            || (package_present && source_length.is_none())
        {
            return Err(error("snapshot package counts differ"));
        }
        let last = self.received.iter().rposition(|count| *count != 0).unwrap();
        for (kind, (&received, expected)) in self
            .received
            .iter()
            .zip(header.chunk_counts)
            .enumerate()
        {
            if received > expected || (kind < last && received != expected) {
                return Err(error("snapshot chunk counts differ"));
            }
        }
        if self.ended && self.received != header.chunk_counts {
            return Err(error("snapshot chunk counts differ"));
        }
        self.authority = Some(AuthorityHydrator::new(&header)?);
        self.model = Some(ModelSnapshotBuilder::new());
        self.preserved = Some(PreservedSnapshotBuilder::new());
        self.facts = package_present.then(PackageFactsBuilder::new);
        self.header = Some(HeaderState {
            header,
            projection_valid,
            graph_present,
            package_present,
            source_length,
            style_counts,
            calculation_counts,
            calculation_list: 0,
            calculation_index: 0,
        });
        Ok(())
    }

    fn accept_calculation(&mut self, payload: &[u8]) -> SnapshotResult<()> {
        let state = self.header.as_mut().expect("validated snapshot header");
        while state.calculation_list < state.calculation_counts.len()
            && state.calculation_index == state.calculation_counts[state.calculation_list]
        {
            state.calculation_list += 1;
            state.calculation_index = 0;
        }
        let mut reader = Reader::new(payload);
        if usize::from(reader.u8()?) != state.calculation_list
            || state.calculation_list >= state.calculation_counts.len()
        {
            return Err(error("snapshot calculation lists are reordered"));
        }
        let address = read_address(&mut reader)?;
        reader.finish()?;
        let CalculationResult {
            changed,
            cycle_cells,
            limited_cells,
        } = &mut state.header.last_calculation;
        match state.calculation_list {
            0 => changed.push(address),
            1 => cycle_cells.push(address),
            2 => limited_cells.push(address),
            _ => return Err(error("invalid snapshot calculation list")),
        }
        state.calculation_index += 1;
        Ok(())
    }

    pub fn advance(&mut self, budget: SnapshotBudget) -> SnapshotResult<SnapshotProgress> {
        if self.failed {
            return Err(error("snapshot builder has failed"));
        }
        if self.ready.is_some() {
            return Ok(SnapshotProgress::ready());
        }
        if let Some(chunk) = self.queue.front()
            && chunk.len() > budget.max_bytes()
        {
            return Err(error("snapshot logical chunk exceeds advance byte budget"));
        }
        if let Some(chunk) = self.queue.front() {
            let (kind, _, payload) = unframe(chunk)?;
            let admission = match kind {
                ChunkKind::AuthorityBase => Some(&mut self.base_admission),
                ChunkKind::Facts => Some(&mut self.facts_admission),
                _ => None,
            };
            if let Some(admission) = admission {
                let mut next = *admission;
                next.push(payload, budget.max_bytes())?;
                *admission = next;
            }
        }
        if self.anchors_done
            && self.graph.is_some()
            && let Some(workbook) = &self.restored
        {
            let mut next = self.graph_cursor;
            if next
                .next_cost(&workbook.model)
                .is_some_and(|bytes| bytes > budget.max_bytes())
            {
                return Err(error("snapshot graph record exceeds advance byte budget"));
            }
            self.graph_cursor = next;
        }
        let result = self.advance_inner(budget);
        self.failed = result.is_err();
        result
    }

    fn advance_inner(&mut self, budget: SnapshotBudget) -> SnapshotResult<SnapshotProgress> {
        if let Some(chunk) = self.queue.pop_front() {
            let (kind, _, payload) = unframe(&chunk)?;
            match kind {
                ChunkKind::Header => {
                    if self.header.is_none() {
                        let snapshot_id = self.snapshot_id.expect("received snapshot identity");
                        self.accept_header(payload, snapshot_id)?;
                    } else {
                        self.accept_calculation(payload)?;
                    }
                }
                ChunkKind::AuthorityBase => {
                    let authority = self.authority.as_mut().expect("validated snapshot header");
                    authority.push_base(payload)?;
                    authority.advance(SnapshotBudget::new(1, budget.max_bytes())?)?;
                }
                ChunkKind::Yrs => self
                    .authority
                    .as_mut()
                    .expect("validated snapshot header")
                    .push_yrs(payload)?,
                ChunkKind::Model | ChunkKind::Cells => self
                    .model
                    .as_mut()
                    .expect("validated snapshot header")
                    .push(&chunk)?,
                ChunkKind::Preserved => self
                    .preserved
                    .as_mut()
                    .expect("validated snapshot header")
                    .push(&chunk)?,
                ChunkKind::Facts => self
                    .facts
                    .as_mut()
                    .ok_or_else(|| error("unexpected snapshot facts"))?
                    .push(payload)
                    .map_err(|failure| error(failure.to_string()))?,
                ChunkKind::Source => {
                    if self.source.is_none() {
                        let length = self
                            .header
                            .as_ref()
                            .and_then(|state| state.source_length)
                            .ok_or_else(|| error("snapshot source length is missing"))?;
                        self.source = Some(SourceContainerBuilder::new(length).map_err(error)?);
                    }
                    self.source
                        .as_mut()
                        .expect("snapshot source builder exists")
                        .push(payload)
                        .map_err(error)?;
                }
                ChunkKind::End => {
                    return Err(error("unexpected queued snapshot control chunk"));
                }
            }
            return Ok(SnapshotProgress::pending());
        }
        if !self.ended || self.fragment.is_some() {
            return Ok(SnapshotProgress::pending());
        }
        if self.restored.is_none() {
            let authority = self.authority.as_mut().expect("validated snapshot header");
            if !authority.advance(budget)?.is_ready() {
                return Ok(SnapshotProgress::pending());
            }
            self.complete_parts()?;
            return Ok(SnapshotProgress::pending());
        }
        let workbook = self.restored.as_mut().expect("completed snapshot parts");
        if !self.anchors_done {
            let mut records = 0;
            let mut bytes = 0usize;
            while records < budget.max_records() {
                let Some(sheet) = workbook.model.sheets.get(self.anchor_sheet) else {
                    self.anchors_done = true;
                    break;
                };
                if workbook.preserved.created[self.anchor_sheet]
                    || workbook.preserved.axes[self.anchor_sheet]
                        .as_ref()
                        .is_some_and(|axes| axes != &xlsx_parse::SheetAxes::default())
                {
                    return Err(error("snapshot preservation state is not initial"));
                }
                if let Some(chart) = sheet.charts.get(self.anchor_chart) {
                    let frame_id = chart.frame_id();
                    let cost = frame_id.len().saturating_add(size_of::<ChartAnchor>());
                    if cost > budget.max_bytes().saturating_sub(bytes) {
                        if records == 0 {
                            return Err(error("snapshot chart exceeds advance byte budget"));
                        }
                        break;
                    }
                    bytes += cost;
                    workbook.opened_anchors.insert(frame_id, chart.anchor);
                    self.anchor_chart += 1;
                } else {
                    self.anchor_sheet += 1;
                    self.anchor_chart = 0;
                }
                records += 1;
            }
            return Ok(SnapshotProgress::pending());
        }
        if let Some(graph) = self.graph.as_mut() {
            if !graph.advance(&workbook.model, NonZeroUsize::MIN) {
                return Ok(SnapshotProgress::pending());
            }
            workbook.graph = self.graph.take().expect("completed graph builder").finish();
        }
        self.ready = Some(HydratedWorkbook {
            workbook: self.restored.take().expect("completed snapshot workbook"),
            calculation_context: self.calculation_context,
        });
        Ok(SnapshotProgress::ready())
    }

    fn complete_parts(&mut self) -> SnapshotResult<()> {
        let HeaderState {
            header,
            projection_valid,
            graph_present,
            package_present,
            source_length: _,
            style_counts,
            calculation_counts,
            calculation_list: _,
            calculation_index: _,
        } = self.header.take().expect("validated snapshot header");
        if self.received != header.chunk_counts {
            return Err(error("snapshot chunk counts differ"));
        }
        if [
            header.last_calculation.changed.len(),
            header.last_calculation.cycle_cells.len(),
            header.last_calculation.limited_cells.len(),
        ] != calculation_counts
        {
            return Err(error("snapshot calculation lists are incomplete"));
        }
        let mut authority = self
            .authority
            .take()
            .expect("snapshot authority builder exists")
            .finish()?;
        authority.set_snapshot_projection_valid(projection_valid);
        let model = self.model.take().expect("snapshot model builder exists").finish()?;
        if model.styles.snapshot_field_counts() != style_counts
            || model.sheets.get(header.active_sheet.0 as usize).is_none()
        {
            return Err(error("snapshot model header differs"));
        }
        let preserved = self
            .preserved
            .take()
            .expect("snapshot preservation builder exists")
            .finish()?;
        let sheets = model.sheets.len();
        if [
            preserved.origins.len(),
            preserved.shared_string_cells.len(),
            preserved.axes.len(),
            preserved.created.len(),
        ]
        .iter()
        .any(|count| *count != sheets)
        {
            return Err(error("snapshot preservation sheet counts differ"));
        }
        let source_container = self
            .source
            .take()
            .map(SourceContainerBuilder::finish)
            .transpose()
            .map_err(error)?;
        let source_package = if package_present {
            let facts = self
                .facts
                .take()
                .expect("snapshot facts builder exists")
                .finish()
                .map_err(|failure| error(failure.to_string()))?;
            let source = source_container
                .as_ref()
                .ok_or_else(|| error("snapshot source container is missing"))?;
            Some(PackageSlot::deferred(source.clone(), facts))
        } else {
            None
        };
        let SnapshotHeader {
            snapshot_id: _,
            mode: _,
            edited_since_open,
            recalculated_since_open,
            moved_references_since_open,
            active_sheet,
            rand_seed,
            model_epoch,
            version_nonce,
            committed_changes,
            last_calculation,
            calculation_context,
            client_id: _,
            guid: _,
            next_sheet_id: _,
            state_vector: _,
            chunk_counts: _,
        } = header;
        self.calculation_context = calculation_context;
        self.graph = graph_present.then(DepGraphBuilder::new);
        self.restored = Some(Workbook {
            authority,
            mode: WorkbookMode::Standalone,
            pending_remote_updates: Vec::new(),
            model,
            source_package,
            source_container,
            preserved,
            preserved_undo: Vec::new(),
            preserved_redo: Vec::new(),
            edited_since_open,
            recalculated_since_open,
            moved_references_since_open,
            active_sheet,
            undo: UndoStack::new(),
            graph: None,
            rand_seed,
            proposals: ProposalSet::new(),
            last_calculation,
            update_observers: Arc::new(Mutex::new(UpdateObservers::default())),
            opened_anchors: BTreeMap::new(),
            sheet_info_cache: Mutex::new(None),
            model_epoch,
            geometry_cache: Mutex::new(HashMap::new()),
            version_nonce,
            committed_changes,
            chart_cache: Mutex::new(ChartCache::default()),
            source_part_hashes: Mutex::new(BTreeMap::new()),
        });
        Ok(())
    }

    pub fn finish(self) -> SnapshotResult<HydratedWorkbook> {
        if self.failed {
            return Err(error("snapshot builder has failed"));
        }
        self.ready.ok_or_else(|| error("snapshot hydration is incomplete"))
    }
}

impl Default for WorkbookSnapshotBuilder {
    fn default() -> Self {
        Self::new()
    }
}

pub struct HydratedWorkbook {
    workbook: Workbook,
    calculation_context: Option<CalculationOptions>,
}

impl HydratedWorkbook {
    pub fn into_parts(self) -> (Workbook, Option<CalculationOptions>) {
        let Self {
            workbook,
            calculation_context,
        } = self;
        (workbook, calculation_context)
    }
}

#[cfg(test)]
#[path = "acceptance_tests.rs"]
mod acceptance_tests;
