use std::collections::LinkedList;
use std::mem::size_of;

use ooxml_opc::{SourceContainer, SourceContainerBuilder, SourceContainerInitializer};
use sha2::{Digest, Sha256};
use xlsx_calc::graph::SnapshotGraphBuilder;
use xlsx_model::CellRef;
use xlsx_parse::{PackageFactsBuilder, PackageFactsEncoder};

use crate::authority::snapshot::{AuthorityHydrator, AuthoritySnapshotEncoder};
use crate::snapshot::header::{SnapshotHeader, SnapshotMode};
use crate::snapshot::model::{ModelSnapshotBuilder, ModelSnapshotEncoder};
use crate::snapshot::preserved::{PreservedSnapshotBuilder, PreservedSnapshotEncoder};
use crate::snapshot::wire::{ChunkKind, Reader, Writer, frame, unframe};
use crate::snapshot::{SnapshotBudget, SnapshotError, SnapshotProgress, SnapshotResult};

use super::{
    Arc, BTreeMap, ChartAnchor, ChartCache, HashMap, Mutex, PackageSlot, PreservedSheetState,
    ProposalSet, SheetId, UndoStack, UpdateObservers, Workbook, WorkbookMode,
};
use crate::{CalculationOptions, CalculationResult, CellAddress};

#[path = "validation.rs"]
mod validation;

fn logical_byte_limit() -> usize {
    crate::snapshot::yrs_split::whole_update_limit() + 12
}

fn record_budget() -> SnapshotResult<SnapshotBudget> {
    SnapshotBudget::new(64, logical_byte_limit())
}

fn index(kind: ChunkKind) -> usize {
    kind as usize - 1
}

fn error(message: impl Into<String>) -> SnapshotError {
    SnapshotError::new(message)
}

#[doc(hidden)]
pub(super) struct Lineage {
    nonce: String,
    changes: u64,
    epoch: u64,
    active_sheet: SheetId,
    seed: Option<u32>,
    recalculated: bool,
    projection_valid: bool,
    state_vector: Vec<u8>,
    authority_revision: u64,
}

impl Lineage {
    fn capture(workbook: &Workbook) -> SnapshotResult<Self> {
        let Workbook {
            authority,
            mode: _,
            pending_remote_updates,
            model,
            source_package,
            snapshot_package_lineage: _,
            source_container,
            preserved,
            preserved_undo,
            preserved_redo,
            edited_since_open,
            recalculated_since_open,
            calculations_since_open,
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
        if *calculations_since_open != u64::from(*recalculated_since_open) {
            return Err(error("snapshot requires an initial calculation count"));
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
        let PreservedSheetState {
            origins,
            shared_string_cells,
            axes,
            created,
        } = preserved;
        let sheets = model.sheets.len();
        if [
            origins.len(),
            shared_string_cells.len(),
            axes.len(),
            created.len(),
        ]
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
            authority_revision: authority.snapshot_revision(),
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
            && self.authority_revision == workbook.authority.snapshot_revision())
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
    retained_package_facts: bool,
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
        let retained_package_facts = workbook
            .snapshot_package_lineage
            .as_ref()
            .map(|lineage| lineage.matches(workbook))
            .transpose()?
            .unwrap_or(false);
        if !retained_package_facts && let Some(package) = &workbook.source_package {
            package
                .materialize()
                .map_err(|failure| error(failure.to_string()))?;
        }
        let lineage = Lineage::capture(workbook)?;
        if budget.max_bytes() < 64 {
            return Err(error("snapshot byte budget is too small for framing"));
        }
        let snapshot_id = u64::from_str_radix(&yrs::uuid_v4().replace('-', "")[..16], 16)
            .map_err(|_| error("cannot allocate snapshot identity"))?;
        let part_bytes = budget
            .max_bytes()
            .saturating_sub(64)
            .clamp(1, logical_byte_limit() - 12);
        let part_budget = SnapshotBudget::new(1, part_bytes)?;
        let record_budget = record_budget()?;
        let authority = AuthoritySnapshotEncoder::new(
            &workbook.authority,
            SnapshotBudget::new(budget.max_records(), part_bytes)?,
        )?;
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
        if let Some(package) = &workbook.source_package {
            let mut facts = PackageFactsEncoder::new();
            while let Some(payload) =
                package.next_facts(&mut facts, part_budget.max_bytes(), retained_package_facts)?
            {
                let ordinal = chunk_counts[index(ChunkKind::Facts)];
                count_chunk(
                    &mut chunk_counts,
                    &frame(ChunkKind::Facts, ordinal, &payload),
                )?;
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
        if header_chunk.len() > logical_byte_limit() {
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
            retained_package_facts,
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
        if self.retained_package_facts
            && !workbook
                .snapshot_package_lineage
                .as_ref()
                .map(|lineage| lineage.matches(workbook))
                .transpose()?
                .unwrap_or(false)
        {
            return Err(error("snapshot retained package permission has changed"));
        }
        if self.pending.is_none() {
            self.pending = self.next_logical(workbook)?;
        }
        let Some(chunk) = &self.pending else {
            return Ok(None);
        };
        let (kind, ordinal, _) = unframe(chunk)?;
        if chunk.len() > logical_byte_limit() {
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
        let record_budget = record_budget()?;
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
                    if let Some(package) = &workbook.source_package
                        && let Some(payload) = package.next_facts(
                            &mut self.facts,
                            self.part_budget.max_bytes(),
                            self.retained_package_facts,
                        )?
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
                            + self
                                .part_budget
                                .max_bytes()
                                .min(bytes.len() - self.source_offset);
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
                    return Ok(Some(frame(
                        ChunkKind::End,
                        0,
                        &self.digest.clone().finalize(),
                    )));
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
    if chunk.len() > logical_byte_limit() {
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
    blocks: BTreeMap<usize, FragmentBlock>,
    received: usize,
}

const FRAGMENT_BLOCK_BYTES: usize = 4096;

#[cfg(test)]
thread_local! {
    static FRAGMENT_ALLOCATED_BYTES: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

struct FragmentBlock {
    bytes: Vec<u8>,
}

impl FragmentBlock {
    fn new(bytes: &[u8]) -> Self {
        let bytes = bytes.to_vec();
        #[cfg(test)]
        FRAGMENT_ALLOCATED_BYTES.set(FRAGMENT_ALLOCATED_BYTES.get() + bytes.capacity());
        Self { bytes }
    }

    fn into_bytes(mut self) -> Vec<u8> {
        #[cfg(test)]
        FRAGMENT_ALLOCATED_BYTES.set(FRAGMENT_ALLOCATED_BYTES.get() - self.bytes.capacity());
        std::mem::take(&mut self.bytes)
    }
}

impl std::ops::Deref for FragmentBlock {
    type Target = [u8];

    fn deref(&self) -> &[u8] {
        &self.bytes
    }
}

impl Drop for FragmentBlock {
    fn drop(&mut self) {
        #[cfg(test)]
        FRAGMENT_ALLOCATED_BYTES.set(FRAGMENT_ALLOCATED_BYTES.get() - self.bytes.capacity());
    }
}

struct QueuedChunk {
    fragment: Option<Fragment>,
    bytes: Vec<u8>,
}

impl std::ops::Deref for QueuedChunk {
    type Target = [u8];

    fn deref(&self) -> &[u8] {
        &self.bytes
    }
}

impl QueuedChunk {
    fn new(mut fragment: Fragment) -> Self {
        let bytes = if fragment.blocks.len() == 1 {
            fragment.blocks.pop_first().unwrap().1.into_bytes()
        } else {
            Vec::new()
        };
        Self {
            fragment: bytes.is_empty().then_some(fragment),
            bytes,
        }
    }

    fn advance(&mut self, budget: SnapshotBudget) -> SnapshotResult<bool> {
        let Some(fragment) = &mut self.fragment else {
            return Ok(true);
        };
        if self.bytes.capacity() == 0 {
            self.bytes
                .try_reserve_exact(fragment.received)
                .map_err(|_| error("cannot allocate received snapshot chunk"))?;
        }
        let mut records = 0;
        let mut copied = 0;
        while records < budget.max_records() && copied < budget.max_bytes() {
            let Some((&offset, block)) = fragment.blocks.first_key_value() else {
                break;
            };
            let start = self.bytes.len() - offset;
            let count = (block.len() - start).min(budget.max_bytes() - copied);
            self.bytes.extend_from_slice(&block[start..start + count]);
            let complete = start + count == block.len();
            copied += count;
            records += 1;
            if complete {
                fragment.blocks.pop_first();
            }
        }
        crate::snapshot::step::record(records, copied);
        if fragment.blocks.is_empty() {
            self.fragment = None;
        }
        Ok(false)
    }
}

#[derive(Clone)]
struct BlockReader<'a> {
    blocks: &'a BTreeMap<usize, FragmentBlock>,
    position: usize,
    end: usize,
}

impl<'a> BlockReader<'a> {
    fn new(fragment: &'a Fragment) -> Self {
        Self {
            blocks: &fragment.blocks,
            position: 0,
            end: fragment.received,
        }
    }

    fn u8(&mut self) -> SnapshotResult<u8> {
        if self.position == self.end {
            return Err(error("short snapshot fragment"));
        }
        let (&offset, block) = self
            .blocks
            .range(..=self.position)
            .next_back()
            .ok_or_else(|| error("snapshot fragment block is missing"))?;
        let byte = block[self.position - offset];
        self.position += 1;
        Ok(byte)
    }

    fn var_u64(&mut self) -> SnapshotResult<u64> {
        let mut value = 0;
        for shift in (0..=63).step_by(7) {
            let byte = self.u8()?;
            if shift == 63 && byte > 1 {
                return Err(error("snapshot integer overflows"));
            }
            value |= u64::from(byte & 0x7f) << shift;
            if byte & 0x80 == 0 {
                return Ok(value);
            }
        }
        Err(error("snapshot integer overflows"))
    }

    fn var_u32(&mut self) -> SnapshotResult<u32> {
        u32::try_from(self.var_u64()?).map_err(|_| error("snapshot integer overflows"))
    }

    fn var_usize(&mut self) -> SnapshotResult<usize> {
        usize::try_from(self.var_u64()?).map_err(|_| error("snapshot integer overflows"))
    }

    fn bool(&mut self) -> SnapshotResult<bool> {
        match self.u8()? {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(error("invalid snapshot boolean")),
        }
    }

    fn option<T>(
        &mut self,
        read: impl FnOnce(&mut Self) -> SnapshotResult<T>,
    ) -> SnapshotResult<Option<T>> {
        if self.bool()? {
            read(self).map(Some)
        } else {
            Ok(None)
        }
    }

    fn f64(&mut self) -> SnapshotResult<f64> {
        let mut bytes = [0; 8];
        for byte in &mut bytes {
            *byte = self.u8()?;
        }
        Ok(f64::from_le_bytes(bytes))
    }

    fn bytes(&mut self) -> SnapshotResult<Self> {
        let length = self.var_usize()?;
        if length > self.end - self.position {
            return Err(error("short snapshot bytes"));
        }
        let mut reader = self.clone();
        reader.end = self.position + length;
        self.position += length;
        Ok(reader)
    }

    fn identity(&mut self, guid: bool) -> SnapshotResult<()> {
        let mut value = self.bytes()?;
        let expected = if guid {
            crate::snapshot::header::GUID_BYTES
        } else {
            crate::snapshot::header::NONCE_BYTES
        };
        if value.end - value.position != expected {
            return Err(error("snapshot identity length is invalid"));
        }
        for index in 0..expected {
            if !crate::snapshot::header::identity_byte(value.u8()?, index, guid) {
                return Err(error("snapshot identity shape is invalid"));
            }
        }
        Ok(())
    }

    fn finish(self) -> SnapshotResult<()> {
        if self.position == self.end {
            Ok(())
        } else {
            Err(error("trailing snapshot bytes"))
        }
    }
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
    calculation_storage: Option<Vec<CellAddress>>,
}

struct HeaderAdmission {
    chunk_counts: [u64; 9],
    source_length: Option<usize>,
}

impl HeaderAdmission {
    fn decode(mut outer: BlockReader<'_>, snapshot_id: u64) -> SnapshotResult<Self> {
        let mut reader = outer.bytes()?;
        if reader.var_u64()? != snapshot_id || reader.u8()? != 0 || reader.bool()? {
            return Err(error("snapshot header is invalid"));
        }
        reader.bool()?;
        if reader.bool()? {
            return Err(error("snapshot header is invalid"));
        }
        reader.var_u32()?;
        reader.option(BlockReader::var_u32)?;
        reader.var_u64()?;
        reader.identity(false)?;
        reader.var_u64()?;
        for _ in 0..3 {
            if reader.var_usize()? != 0 {
                return Err(error("snapshot calculation counts differ"));
            }
        }
        reader.option(|reader| reader.option(BlockReader::f64))?;
        reader.var_u64()?;
        reader.identity(true)?;
        if reader.var_u64()? != 0 {
            return Err(error("snapshot header is invalid"));
        }
        reader.bytes()?;
        let mut chunk_counts = [0; 9];
        for count in &mut chunk_counts {
            *count = reader.var_u64()?;
        }
        reader.finish()?;
        outer.bool()?;
        outer.bool()?;
        let package_present = outer.bool()?;
        let source_length = outer.option(BlockReader::var_usize)?;
        for _ in 0..7 {
            outer.var_usize()?;
        }
        let mut header_count = 1u64;
        for _ in 0..3 {
            header_count = header_count
                .checked_add(outer.var_usize()? as u64)
                .ok_or_else(|| error("snapshot header count overflows"))?;
        }
        outer.finish()?;
        if header_count != chunk_counts[index(ChunkKind::Header)] {
            return Err(error("snapshot calculation counts differ"));
        }
        if chunk_counts[index(ChunkKind::End)] != 1
            || chunk_counts[index(ChunkKind::Model)] == 0
            || chunk_counts[index(ChunkKind::Preserved)] == 0
        {
            return Err(error("snapshot header is invalid"));
        }
        if package_present != (chunk_counts[index(ChunkKind::Facts)] != 0)
            || source_length.is_some() != (chunk_counts[index(ChunkKind::Source)] != 0)
            || (package_present && source_length.is_none())
        {
            return Err(error("snapshot package counts differ"));
        }
        Ok(Self {
            chunk_counts,
            source_length,
        })
    }
}

enum SourceBuilder {
    Initializing(SourceContainerInitializer),
    Ready(SourceContainerBuilder),
}

impl SourceBuilder {
    fn finish(self) -> Result<SourceContainer, String> {
        match self {
            Self::Initializing(_) => Err("source initialization is incomplete".to_owned()),
            Self::Ready(builder) => builder.finish(),
        }
    }
}

#[derive(Clone, Copy, Default)]
struct StreamAdmission {
    remaining: usize,
    length: usize,
    varint: u64,
    shift: u32,
    records: usize,
    bytes: usize,
}

impl StreamAdmission {
    fn push(
        &mut self,
        mut payload: &[u8],
        budget: SnapshotBudget,
        facts: bool,
    ) -> SnapshotResult<()> {
        let max_bytes = budget.max_bytes();
        self.records = 0;
        self.bytes = if facts { 0 } else { payload.len() };
        let record_limit =
            if facts { max_bytes } else { usize::MAX }.min(xlsx_parse::SNAPSHOT_RECORD_MAX_BYTES);
        if self.remaining != 0 && self.length > record_limit {
            return Err(error("snapshot record exceeds advance byte budget"));
        }
        while !payload.is_empty() {
            if self.remaining != 0 {
                let count = self.remaining.min(payload.len());
                self.remaining -= count;
                payload = &payload[count..];
                if self.remaining == 0 {
                    self.records += 1;
                    if facts {
                        self.bytes = self.bytes.saturating_add(self.length);
                    }
                    if self.records > budget.max_records() || self.bytes > max_bytes {
                        return Err(error(
                            "snapshot stream exceeds advance record or byte budget",
                        ));
                    }
                }
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
                    if length > record_limit {
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

#[doc(hidden)]
pub struct WorkbookSnapshotBuilder {
    max_frame_bytes: usize,
    header: Option<HeaderState>,
    admission: Option<HeaderAdmission>,
    snapshot_id: Option<u64>,
    ordinals: [u64; 9],
    received: [u64; 9],
    fragment: Option<Fragment>,
    queue: LinkedList<QueuedChunk>,
    digest: Sha256,
    ended: bool,
    failed: bool,
    authority: Option<AuthorityHydrator>,
    model: Option<ModelSnapshotBuilder>,
    preserved: Option<PreservedSnapshotBuilder>,
    facts: Option<PackageFactsBuilder>,
    source: Option<SourceBuilder>,
    source_received_bytes: usize,
    restored: Option<Workbook>,
    graph: Option<SnapshotGraphBuilder>,
    anchor_sheet: usize,
    anchor_chart: usize,
    anchors_done: bool,
    validation: validation::ModelValidation,
    validated: bool,
    authority_validated: bool,
    authority_validation: crate::authority::snapshot_validation::SnapshotValidation,
    authority_keys: crate::snapshot::yrs_split::SnapshotKeys,
    ready: Option<HydratedWorkbook>,
    calculation_context: Option<CalculationOptions>,
    base_admission: StreamAdmission,
    facts_admission: StreamAdmission,
}

impl WorkbookSnapshotBuilder {
    pub fn new() -> Self {
        Self {
            max_frame_bytes: logical_byte_limit().saturating_add(64),
            header: None,
            admission: None,
            snapshot_id: None,
            ordinals: [0; 9],
            received: [0; 9],
            fragment: None,
            queue: LinkedList::new(),
            digest: Sha256::new(),
            ended: false,
            failed: false,
            authority: None,
            model: None,
            preserved: None,
            facts: None,
            source: None,
            source_received_bytes: 0,
            restored: None,
            graph: None,
            anchor_sheet: 0,
            anchor_chart: 0,
            anchors_done: false,
            validation: validation::ModelValidation::default(),
            validated: false,
            authority_validated: false,
            authority_validation:
                crate::authority::snapshot_validation::SnapshotValidation::default(),
            authority_keys: crate::snapshot::yrs_split::SnapshotKeys::default(),
            ready: None,
            calculation_context: None,
            base_admission: StreamAdmission::default(),
            facts_admission: StreamAdmission::default(),
        }
    }

    pub fn with_max_frame_bytes(max_frame_bytes: usize) -> SnapshotResult<Self> {
        if max_frame_bytes < 64 {
            return Err(error("snapshot frame limit is too small for framing"));
        }
        Ok(Self {
            max_frame_bytes,
            ..Self::new()
        })
    }

    pub fn push(&mut self, chunk: &[u8]) -> SnapshotResult<SnapshotProgress> {
        crate::snapshot::step::reset();
        if self.failed {
            return Err(error("snapshot builder has failed"));
        }
        let result = self.push_inner(chunk);
        self.failed = result.is_err();
        result
    }

    fn push_inner(&mut self, chunk: &[u8]) -> SnapshotResult<SnapshotProgress> {
        if chunk.len() > self.max_frame_bytes {
            return Err(error(
                "snapshot transport frame exceeds accepted byte limit",
            ));
        }
        if self.ended || self.ready.is_some() {
            return Err(error("snapshot has extra chunks"));
        }
        let (kind, ordinal, payload) = unframe(chunk)?;
        crate::snapshot::step::record(0, chunk.len());
        if ordinal != self.ordinals[index(kind)] {
            return Err(error("snapshot transport ordinal differs"));
        }
        let mut reader = Reader::new(payload);
        let snapshot_id = reader.var_u64()?;
        let logical_ordinal = reader.var_u64()?;
        let length = reader.var_usize()?;
        let offset = reader.var_usize()?;
        let bytes = reader.rest();
        if self
            .snapshot_id
            .is_some_and(|expected| expected != snapshot_id)
        {
            return Err(error("snapshot lineage differs"));
        }
        if bytes.is_empty() || length == 0 || length > logical_byte_limit() {
            return Err(error("invalid snapshot fragment length"));
        }
        if self.fragment.is_none() {
            let expected_kind = match &self.admission {
                None => index(ChunkKind::Header),
                Some(state) => self
                    .received
                    .iter()
                    .zip(state.chunk_counts)
                    .position(|(received, expected)| *received < expected)
                    .ok_or_else(|| error("snapshot has extra chunks"))?,
            };
            if index(kind) != expected_kind || offset != 0 {
                return Err(error("snapshot chunks are missing or reordered"));
            }
            let expected_ordinal = if kind == ChunkKind::Cells {
                self.received[index(ChunkKind::Model)] + self.received[index(ChunkKind::Cells)]
            } else {
                self.received[index(kind)]
            };
            if logical_ordinal != expected_ordinal {
                return Err(error("snapshot logical ordinal differs"));
            }
            self.fragment = Some(Fragment {
                kind,
                ordinal: logical_ordinal,
                length,
                blocks: BTreeMap::new(),
                received: 0,
            });
        }
        let fragment = self
            .fragment
            .as_mut()
            .ok_or_else(|| error("snapshot fragment is missing"))?;
        if fragment.kind != kind
            || fragment.ordinal != logical_ordinal
            || fragment.length != length
            || fragment.received != offset
            || bytes.len() > length - offset
        {
            return Err(error("snapshot fragments are missing or reordered"));
        }
        #[cfg(test)]
        crate::snapshot::step::allocate(bytes.len());
        for block in bytes.chunks(FRAGMENT_BLOCK_BYTES) {
            fragment
                .blocks
                .insert(fragment.received, FragmentBlock::new(block));
            fragment.received += block.len();
        }
        if kind != ChunkKind::End {
            self.digest.update(bytes);
        }
        self.snapshot_id = Some(snapshot_id);
        self.ordinals[index(kind)] = ordinal
            .checked_add(1)
            .ok_or_else(|| error("snapshot ordinal overflows"))?;
        if fragment.received != length {
            return Ok(SnapshotProgress::pending());
        }
        let fragment = self
            .fragment
            .take()
            .ok_or_else(|| error("snapshot fragment is missing"))?;
        let mut payload = BlockReader::new(&fragment);
        let version = payload.u8()?;
        if version != crate::snapshot::wire::FORMAT_VERSION
            && version != crate::snapshot::wire::PACKED_FORMAT_VERSION
        {
            return Err(error("unsupported snapshot format"));
        }
        let inner_kind = payload.u8()?;
        let inner_ordinal = payload.var_u64()?;
        if version == crate::snapshot::wire::PACKED_FORMAT_VERSION {
            let count = payload.var_usize()?;
            if !matches!(
                kind,
                ChunkKind::Model | ChunkKind::Cells | ChunkKind::Preserved
            ) || !(2..=64).contains(&count)
            {
                return Err(error("invalid packed snapshot chunk"));
            }
            let mut length = 0usize;
            for _ in 0..count {
                let bytes = payload.var_usize()?;
                if bytes == 0 {
                    return Err(error("empty packed snapshot record"));
                }
                length = length
                    .checked_add(bytes)
                    .ok_or_else(|| error("packed snapshot length overflows"))?;
            }
            if length != payload.end - payload.position {
                return Err(error("packed snapshot length differs"));
            }
        }
        if let Some(state) = &self.admission {
            let expected = self
                .received
                .iter()
                .zip(state.chunk_counts)
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
        if inner_kind != kind as u8
            || inner_ordinal != logical_ordinal
            || inner_ordinal != expected_ordinal
        {
            return Err(error("snapshot logical ordinal differs"));
        }
        match kind {
            ChunkKind::End => {
                let admission = self
                    .admission
                    .as_ref()
                    .ok_or_else(|| error("snapshot header is missing"))?;
                if admission.source_length.unwrap_or(0) != self.source_received_bytes {
                    return Err(error("snapshot source length differs"));
                }
                let actual = self.digest.clone().finalize();
                for byte in actual {
                    if payload.u8()? != byte {
                        return Err(error("snapshot content digest differs"));
                    }
                }
                payload.finish()?;
                self.ended = true;
                self.queue.push_back(QueuedChunk::new(fragment));
            }
            _ => {
                if kind == ChunkKind::Header && inner_ordinal == 0 {
                    self.admission = Some(HeaderAdmission::decode(payload.clone(), snapshot_id)?);
                }
                if kind == ChunkKind::Source {
                    self.source_received_bytes = self
                        .source_received_bytes
                        .checked_add(payload.end - payload.position)
                        .ok_or_else(|| error("snapshot source length overflows"))?;
                    let expected = self
                        .admission
                        .as_ref()
                        .and_then(|state| state.source_length)
                        .ok_or_else(|| error("snapshot source length is missing"))?;
                    if self.source_received_bytes > expected {
                        return Err(error("snapshot source length differs"));
                    }
                }
                self.queue.push_back(QueuedChunk::new(fragment));
            }
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
        if package_present != (header.chunk_count(ChunkKind::Facts) != 0)
            || source_length.is_some() != (header.chunk_count(ChunkKind::Source) != 0)
            || (package_present && source_length.is_none())
        {
            return Err(error("snapshot package counts differ"));
        }
        let last = self
            .received
            .iter()
            .rposition(|count| *count != 0)
            .ok_or_else(|| error("snapshot header is missing"))?;
        for (kind, (&received, expected)) in
            self.received.iter().zip(header.chunk_counts).enumerate()
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
            calculation_storage: None,
        });
        Ok(())
    }

    fn accept_calculation(&mut self, payload: &[u8]) -> SnapshotResult<()> {
        let state = self
            .header
            .as_mut()
            .ok_or_else(|| error("snapshot header is missing"))?;
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

    fn advance_calculation_capacity(&mut self, budget: SnapshotBudget) -> SnapshotResult<bool> {
        let state = self
            .header
            .as_mut()
            .ok_or_else(|| error("snapshot header is missing"))?;
        while state.calculation_list < state.calculation_counts.len()
            && state.calculation_index == state.calculation_counts[state.calculation_list]
        {
            state.calculation_list += 1;
            state.calculation_index = 0;
        }
        let CalculationResult {
            changed,
            cycle_cells,
            limited_cells,
        } = &mut state.header.last_calculation;
        let list = match state.calculation_list {
            0 => changed,
            1 => cycle_cells,
            2 => limited_cells,
            _ => return Err(error("invalid snapshot calculation list")),
        };
        if state.calculation_storage.is_none() {
            if list.len() < list.capacity() {
                return Ok(true);
            }
            let capacity = list.len().saturating_mul(2).max(16);
            let mut storage = Vec::new();
            storage
                .try_reserve_exact(capacity)
                .map_err(|_| error("cannot allocate snapshot calculation"))?;
            if list.is_empty() {
                *list = storage;
                return Ok(true);
            }
            state.calculation_storage = Some(storage);
        }
        let storage = state
            .calculation_storage
            .as_mut()
            .ok_or_else(|| error("snapshot calculation storage is missing"))?;
        let mut records = 0;
        let mut bytes = 0;
        while let Some(address) = list.get(storage.len()) {
            if records == budget.max_records()
                || size_of::<CellAddress>() > budget.max_bytes().saturating_sub(bytes)
            {
                if records == 0 {
                    return Err(error("snapshot calculation exceeds advance byte budget"));
                }
                break;
            }
            storage.push(*address);
            records += 1;
            bytes += size_of::<CellAddress>();
        }
        crate::snapshot::step::record(records, bytes);
        if storage.len() == list.len() {
            *list = state
                .calculation_storage
                .take()
                .ok_or_else(|| error("snapshot calculation storage is missing"))?;
        }
        Ok(false)
    }

    pub fn advance(&mut self, budget: SnapshotBudget) -> SnapshotResult<SnapshotProgress> {
        crate::snapshot::step::reset();
        if self.failed {
            return Err(error("snapshot builder has failed"));
        }
        if self.ready.is_some() {
            return Ok(SnapshotProgress::ready());
        }
        if budget.max_bytes() < size_of::<xlsx_model::Sheet>() {
            return Err(error("snapshot storage exceeds advance byte budget"));
        }
        loop {
            let work = crate::snapshot::step::current();
            let Some(remaining) = budget.remaining(work.records, work.bytes) else {
                return Ok(SnapshotProgress::pending());
            };
            let state = self.drain_state();
            match self.advance_unit(remaining) {
                Ok(progress) if progress.is_ready() => return Ok(progress),
                Ok(_) => {}
                Err(failure)
                    if remaining.is_partial()
                        && (failure.is_budget_refusal()
                            || failure.to_string()
                                == "snapshot facts chunk completes multiple records") =>
                {
                    return Ok(SnapshotProgress::pending());
                }
                Err(failure) => return Err(failure),
            }
            let next = crate::snapshot::step::current();
            if next.records == work.records
                && next.bytes == work.bytes
                && self.drain_state() == state
                && (!self.queue.is_empty() || !self.ended || self.fragment.is_some())
            {
                return Ok(SnapshotProgress::pending());
            }
        }
    }

    fn drain_state(&self) -> (usize, bool, bool, bool, bool, bool, bool, bool) {
        (
            self.queue.len(),
            self.header.is_some(),
            self.restored.is_some(),
            self.validated,
            self.authority_validated,
            self.anchors_done,
            self.graph.is_some(),
            matches!(self.source, Some(SourceBuilder::Ready(_))),
        )
    }

    fn advance_unit(&mut self, budget: SnapshotBudget) -> SnapshotResult<SnapshotProgress> {
        if let Some(chunk) = self.queue.front_mut()
            && !chunk.advance(budget)?
        {
            return Ok(SnapshotProgress::pending());
        }
        if let Some(chunk) = self.queue.front()
            && chunk.len() > budget.max_bytes()
            && !matches!(unframe(chunk)?.0, ChunkKind::Yrs | ChunkKind::Model)
            && chunk.first() != Some(&crate::snapshot::wire::PACKED_FORMAT_VERSION)
        {
            return Err(error("snapshot logical chunk exceeds advance byte budget"));
        }
        if let Some(authority) = &mut self.authority
            && authority.has_pending_base()
        {
            return authority
                .advance_bounded(budget)
                .map(|_| SnapshotProgress::pending());
        }
        let calculation = self.header.is_some()
            && self
                .queue
                .front()
                .map(|chunk| unframe(chunk).map(|(kind, _, _)| kind == ChunkKind::Header))
                .transpose()?
                .unwrap_or(false);
        if calculation && !self.advance_calculation_capacity(budget)? {
            return Ok(SnapshotProgress::pending());
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
                next.push(payload, budget, kind == ChunkKind::Facts)?;
                if kind == ChunkKind::Facts && next.records > 1 {
                    self.failed = !budget.is_partial();
                    return Err(error("snapshot facts chunk completes multiple records"));
                }
            }
        }
        if self
            .queue
            .front()
            .is_some_and(|chunk| unframe(chunk).is_ok_and(|(kind, _, _)| kind == ChunkKind::Facts))
        {
            let (ready, records, bytes) = self
                .facts
                .as_mut()
                .ok_or_else(|| error("snapshot facts builder is missing"))?
                .advance_capacity(budget.max_records(), budget.max_bytes())
                .map_err(|failure| error(failure.to_string()))?;
            crate::snapshot::step::record(records, bytes);
            if !ready {
                return Ok(SnapshotProgress::pending());
            }
        }
        if let Some(chunk) = self.queue.front() {
            let (kind, _, payload) = unframe(chunk)?;
            let admission = match kind {
                ChunkKind::AuthorityBase => Some(&mut self.base_admission),
                ChunkKind::Facts => Some(&mut self.facts_admission),
                _ => None,
            };
            if let Some(admission) = admission {
                admission.push(payload, budget, kind == ChunkKind::Facts)?;
            }
        }
        let result = self.advance_inner(budget);
        self.failed = result.as_ref().err().is_some_and(|failure| {
            !(budget.is_partial() && failure.is_budget_refusal())
                && !matches!(
                    failure.to_string().as_str(),
                    "snapshot graph record exceeds advance byte budget"
                        | "snapshot decoding exceeds advance byte budget"
                        | "snapshot authority exceeds advance byte budget"
                        | "Yrs snapshot record exceeds advance byte budget"
                        | "snapshot authority validation exceeds advance byte budget"
                        | "snapshot authority retirement exceeds advance byte budget"
                )
        });
        result
    }

    fn advance_inner(&mut self, budget: SnapshotBudget) -> SnapshotResult<SnapshotProgress> {
        if let Some(chunk) = self.queue.front() {
            let (kind, ordinal, payload) = unframe(chunk)?;
            if kind == ChunkKind::Yrs {
                let authority = self
                    .authority
                    .as_mut()
                    .ok_or_else(|| error("snapshot header is missing"))?;
                if authority.advance_yrs(payload, budget)?.is_ready() {
                    self.queue.pop_front();
                }
                return Ok(SnapshotProgress::pending());
            }
            if matches!(
                kind,
                ChunkKind::Model | ChunkKind::Cells | ChunkKind::Preserved
            ) {
                let progress = if kind == ChunkKind::Preserved {
                    self.preserved
                        .as_mut()
                        .ok_or_else(|| error("snapshot header is missing"))?
                        .advance_bounded(chunk, budget)?
                } else {
                    self.model
                        .as_mut()
                        .ok_or_else(|| error("snapshot header is missing"))?
                        .advance_bounded(chunk, budget)?
                };
                if progress.is_ready() {
                    let header = &self
                        .header
                        .as_ref()
                        .ok_or_else(|| error("snapshot header is missing"))?
                        .header;
                    let next = ordinal
                        .checked_add(1)
                        .ok_or_else(|| error("snapshot ordinal overflows"))?;
                    if kind == ChunkKind::Preserved {
                        if next == header.chunk_count(ChunkKind::Preserved) {
                            self.preserved.as_ref().unwrap().validate_complete()?;
                        }
                    } else if next
                        == header
                            .chunk_count(ChunkKind::Model)
                            .checked_add(header.chunk_count(ChunkKind::Cells))
                            .ok_or_else(|| error("snapshot model count overflows"))?
                    {
                        self.model.as_ref().unwrap().validate_complete()?;
                    }
                    self.queue.pop_front();
                }
                return Ok(SnapshotProgress::pending());
            }
            if kind == ChunkKind::Source {
                if !self.ended {
                    return Ok(SnapshotProgress::pending());
                }
                if self.source.is_none() {
                    let length = self
                        .header
                        .as_ref()
                        .and_then(|state| state.source_length)
                        .ok_or_else(|| error("snapshot source length is missing"))?;
                    self.source = Some(SourceBuilder::Initializing(
                        SourceContainerInitializer::new(length).map_err(error)?,
                    ));
                }
                if let Some(SourceBuilder::Initializing(initializer)) = self.source.as_mut() {
                    let _count = initializer.advance(budget.max_bytes()).map_err(error)?;
                    crate::snapshot::step::initialize(_count);
                    if initializer.is_ready() {
                        let Some(SourceBuilder::Initializing(initializer)) = self.source.take()
                        else {
                            return Err(error("snapshot source initializer is missing"));
                        };
                        self.source =
                            Some(SourceBuilder::Ready(initializer.finish().map_err(error)?));
                    }
                    return Ok(SnapshotProgress::pending());
                }
            }
        }
        if let Some(chunk) = self.queue.pop_front() {
            let (kind, _, payload) = unframe(&chunk)?;
            match kind {
                ChunkKind::Header => {
                    if self.header.is_none() {
                        let snapshot_id = self
                            .snapshot_id
                            .ok_or_else(|| error("snapshot identity is missing"))?;
                        self.accept_header(payload, snapshot_id)?;
                    } else {
                        self.accept_calculation(payload)?;
                    }
                    crate::snapshot::step::record(1, chunk.len());
                }
                ChunkKind::AuthorityBase => {
                    let authority = self
                        .authority
                        .as_mut()
                        .ok_or_else(|| error("snapshot header is missing"))?;
                    authority.push_base(payload)?;
                    crate::snapshot::step::record(1, payload.len());
                }
                ChunkKind::Yrs => return Err(error("unexpected queued Yrs chunk")),
                ChunkKind::Model | ChunkKind::Cells | ChunkKind::Preserved => {
                    return Err(error("unexpected queued snapshot metadata chunk"));
                }
                ChunkKind::Facts => {
                    self.facts
                        .as_mut()
                        .ok_or_else(|| error("unexpected snapshot facts"))?
                        .push(payload)
                        .map_err(|failure| error(failure.to_string()))?;
                    crate::snapshot::step::record(
                        self.facts_admission.records,
                        self.facts_admission.bytes.max(payload.len()),
                    );
                }
                ChunkKind::Source => {
                    let Some(SourceBuilder::Ready(source)) = self.source.as_mut() else {
                        return Err(error("snapshot source builder is missing"));
                    };
                    source.push(payload).map_err(error)?;
                    crate::snapshot::step::record(1, payload.len());
                }
                ChunkKind::End => {
                    crate::snapshot::step::record(1, chunk.len());
                }
            }
            return Ok(SnapshotProgress::pending());
        }
        if !self.ended || self.fragment.is_some() {
            return Ok(SnapshotProgress::pending());
        }
        if self.restored.is_none() {
            let authority = self
                .authority
                .as_mut()
                .ok_or_else(|| error("snapshot header is missing"))?;
            if !authority.advance_bounded(budget)?.is_ready() {
                return Ok(SnapshotProgress::pending());
            }
            if !authority.advance_finalization(budget)?.is_ready() {
                return Ok(SnapshotProgress::pending());
            }
            self.authority_keys = authority.take_validation_keys();
            self.complete_parts()?;
            return Ok(SnapshotProgress::pending());
        }
        let workbook = self
            .restored
            .as_mut()
            .ok_or_else(|| error("snapshot parts are missing"))?;
        if !self.validated {
            let (ready, _, _) = self.validation.advance(
                &workbook.model,
                workbook.source_package.is_some(),
                budget,
            )?;
            self.validated = ready;
            return Ok(SnapshotProgress::pending());
        }
        if !self.authority_validated {
            self.authority_validated = workbook.authority.validate_snapshot_model(
                &workbook.model,
                &mut self.authority_validation,
                &mut self.authority_keys,
                budget,
            )?;
            return Ok(SnapshotProgress::pending());
        }
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
                    let digits = chart
                        .anchor_index
                        .checked_ilog10()
                        .map_or(1, |log| log as usize + 1);
                    let cost = chart
                        .drawing
                        .len()
                        .saturating_add(1 + digits)
                        .saturating_add(size_of::<ChartAnchor>());
                    if cost > budget.max_bytes().saturating_sub(bytes) {
                        if records == 0 {
                            return Err(error("snapshot chart exceeds advance byte budget"));
                        }
                        break;
                    }
                    bytes += cost;
                    #[cfg(test)]
                    crate::snapshot::step::allocate(cost);
                    workbook
                        .opened_anchors
                        .insert(chart.frame_id(), chart.anchor);
                    self.anchor_chart += 1;
                } else {
                    self.anchor_sheet += 1;
                    self.anchor_chart = 0;
                }
                records += 1;
            }
            crate::snapshot::step::record(records, bytes);
            return Ok(SnapshotProgress::pending());
        }
        if let Some(graph) = self.graph.as_mut() {
            #[cfg(test)]
            let migrated = graph.migrated_entries();
            let mut ready = false;
            let mut _records = 0;
            let mut bytes = 0;
            for visit in 0..budget.max_records() {
                match graph.advance(&workbook.model, budget.max_bytes() - bytes) {
                    Ok((done, count, cost)) => {
                        _records += count;
                        bytes += cost;
                        ready = done;
                        if ready {
                            break;
                        }
                    }
                    Err(failure)
                        if visit != 0
                            && failure == "snapshot graph record exceeds advance byte budget" =>
                    {
                        break;
                    }
                    Err(failure) => return Err(error(failure)),
                }
            }
            crate::snapshot::step::record(_records, bytes);
            #[cfg(test)]
            crate::snapshot::step::migrate(graph.migrated_entries() - migrated);
            if !ready {
                return Ok(SnapshotProgress::pending());
            }
            workbook.graph = self
                .graph
                .take()
                .ok_or_else(|| error("completed graph builder is missing"))?
                .finish();
        }
        self.ready = Some(HydratedWorkbook {
            workbook: self
                .restored
                .take()
                .ok_or_else(|| error("completed snapshot workbook is missing"))?,
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
            calculation_storage: _,
        } = self
            .header
            .take()
            .ok_or_else(|| error("validated snapshot header is missing"))?;
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
            .ok_or_else(|| error("snapshot authority builder is missing"))?
            .finish_drained()?;
        authority.set_snapshot_projection_valid(projection_valid);
        let model = self
            .model
            .take()
            .ok_or_else(|| error("snapshot model builder is missing"))?
            .finish()?;
        if model.styles.snapshot_field_counts() != style_counts
            || model.sheets.get(header.active_sheet.0 as usize).is_none()
        {
            return Err(error("snapshot model header differs"));
        }
        let preserved = self
            .preserved
            .take()
            .ok_or_else(|| error("snapshot preservation builder is missing"))?
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
            .map(SourceBuilder::finish)
            .transpose()
            .map_err(error)?;
        let source_package = if package_present {
            let facts = self
                .facts
                .take()
                .ok_or_else(|| error("snapshot facts builder is missing"))?
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
            state_vector,
            chunk_counts: _,
        } = header;
        self.calculation_context = calculation_context;
        self.graph = graph_present.then(SnapshotGraphBuilder::new);
        let snapshot_package_lineage = Some(Lineage {
            nonce: version_nonce.clone(),
            changes: committed_changes,
            epoch: model_epoch,
            active_sheet,
            seed: rand_seed,
            recalculated: recalculated_since_open,
            projection_valid,
            state_vector,
            authority_revision: authority.snapshot_revision(),
        });
        self.restored = Some(Workbook {
            authority,
            mode: WorkbookMode::Standalone,
            pending_remote_updates: Vec::new(),
            model,
            source_package,
            snapshot_package_lineage,
            source_container,
            preserved,
            preserved_undo: Vec::new(),
            preserved_redo: Vec::new(),
            edited_since_open,
            recalculated_since_open,
            calculations_since_open: u64::from(recalculated_since_open),
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
        self.ready
            .ok_or_else(|| error("snapshot hydration is incomplete"))
    }
}

impl Default for WorkbookSnapshotBuilder {
    fn default() -> Self {
        Self::new()
    }
}

#[doc(hidden)]
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
