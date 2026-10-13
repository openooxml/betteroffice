use crate::authority::SheetOrderHistory;
use crate::snapshot::preserved::{PreservedSnapshotBuilder, PreservedSnapshotEncoder};
use crate::snapshot::wire::{Reader, Writer};
use crate::snapshot::{SnapshotBudget, SnapshotResult};
use sha2::{Digest, Sha256};

use super::*;

const MAGIC: &[u8; 8] = b"XLSXRECV";
const CODEC_VERSION: u8 = 3;

#[derive(Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct RecoveryCheckpoint {
    engine: String,
    through_seq: u64,
    header: RecoveryHeader,
    calculation: RecoveryCalculation,
    collaborative: bool,
    proposals: Vec<RecoveryProposal>,
    proposal_id_counter: u64,
    axes: Vec<Option<RecoveryAxes>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    history: Option<RecoveryHistory>,
}

#[derive(Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct RecoveryHistory {
    undo: UndoStack,
    sheet_order: SheetOrderHistory,
    preserved_undo: Vec<RecoveryPreservedHistory>,
    preserved_redo: Vec<RecoveryPreservedHistory>,
}

#[derive(Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct RecoveryPreservedHistory {
    before: RecoveryPreserved,
    after: RecoveryPreserved,
}

#[derive(Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct RecoveryPreserved {
    chunks: Vec<Vec<u8>>,
    axes: Vec<Option<RecoveryAxes>>,
}

impl RecoveryPreserved {
    fn capture(state: &PreservedSheetState) -> Result<Self> {
        let mut markers = state.clone();
        markers.axes = state
            .axes
            .iter()
            .map(|axes| axes.as_ref().map(|_| xlsx_parse::SheetAxes::default()))
            .collect();
        let mut encoder = PreservedSnapshotEncoder::new();
        let mut chunks = Vec::new();
        let budget = budget().map_err(invalid)?;
        while let Some(chunk) = encoder.next(&markers, budget).map_err(invalid)? {
            chunks.push(chunk);
        }
        Ok(Self {
            chunks,
            axes: state
                .axes
                .iter()
                .map(|axes| axes.as_ref().map(RecoveryAxes::capture).transpose())
                .collect::<Result<_>>()?,
        })
    }

    fn restore(self) -> Result<PreservedSheetState> {
        let mut builder = PreservedSnapshotBuilder::new();
        let budget = budget().map_err(invalid)?;
        for chunk in &self.chunks {
            while !builder
                .advance_bounded(chunk, budget)
                .map_err(invalid)?
                .is_ready()
            {}
        }
        let mut state = builder.finish().map_err(invalid)?;
        if self.axes.len() != state.axes.len()
            || self
                .axes
                .iter()
                .zip(&state.axes)
                .any(|(axes, marker)| axes.is_some() != marker.is_some())
            || [
                state.origins.len(),
                state.shared_string_cells.len(),
                state.created.len(),
            ]
            .iter()
            .any(|length| *length != state.axes.len())
        {
            return Err(invalid("Recovery history sheet count mismatch"));
        }
        state.axes = self
            .axes
            .into_iter()
            .map(|axes| axes.map(RecoveryAxes::restore).transpose())
            .collect::<Result<_>>()?;
        Ok(state)
    }
}

impl RecoveryPreservedHistory {
    fn capture(history: &PreservedStateHistory) -> Result<Self> {
        Ok(Self {
            before: RecoveryPreserved::capture(&history.before)?,
            after: RecoveryPreserved::capture(&history.after)?,
        })
    }

    fn restore(self) -> Result<PreservedStateHistory> {
        Ok(PreservedStateHistory {
            before: self.before.restore()?,
            after: self.after.restore()?,
        })
    }
}

#[derive(Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct RecoveryHeader {
    edited_since_open: bool,
    recalculated_since_open: bool,
    moved_references_since_open: bool,
    active_sheet: SheetId,
    rand_seed: Option<u32>,
    version_nonce: String,
    committed_changes: u64,
    last_calculation: RecoveryDiagnostics,
    calculation_context: Option<CalculationOptions>,
    client_id: u64,
    next_sheet_id: u64,
}

#[derive(Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct RecoveryDiagnostics {
    cycle_cells: Vec<CellAddress>,
    limited_cells: Vec<CellAddress>,
}

#[derive(Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct RecoveryAxes {
    rows: Vec<(u32, u32, u32)>,
    cols: Vec<(u32, u32, u32)>,
}

impl RecoveryAxes {
    fn capture(axes: &xlsx_parse::SheetAxes) -> Result<Self> {
        fn segments(
            ranges: Vec<std::ops::Range<u32>>,
            source: impl Fn(u32) -> Option<u32>,
        ) -> Result<Vec<(u32, u32, u32)>> {
            ranges
                .into_iter()
                .map(|range| {
                    let start = source(range.start)
                        .ok_or_else(|| invalid("Recovery axis has no source"))?;
                    Ok((start, range.start, range.end - range.start))
                })
                .collect()
        }
        Ok(Self {
            rows: segments(axes.rows.current_ranges(0..MAX_ROWS), |at| {
                axes.rows.source(at)
            })?,
            cols: segments(axes.cols.current_ranges(0..MAX_COLS), |at| {
                axes.cols.source(at)
            })?,
        })
    }

    fn restore(self) -> Result<xlsx_parse::SheetAxes> {
        fn restore_axis(
            segments: &[(u32, u32, u32)],
            limit: u32,
            mut edit: impl FnMut(bool, u32, u32),
        ) -> Result<()> {
            let mut source_end = 0;
            let mut current_end = 0;
            for &(source, current, len) in segments {
                if len == 0
                    || source < source_end
                    || current < current_end
                    || source >= limit
                    || current >= limit
                    || len > limit - source
                    || len > limit - current
                {
                    return Err(invalid("Invalid recovery axis segment"));
                }
                source_end = source + len;
                current_end = current + len;
            }
            let mut source_end = 0;
            let mut cursor = 0;
            for &(source, _, len) in segments {
                if source > source_end {
                    edit(false, cursor, source - source_end);
                }
                cursor += len;
                source_end = source + len;
            }
            if source_end < limit {
                edit(false, cursor, limit - source_end);
            }
            let mut cursor = 0;
            for &(_, current, len) in segments {
                if current > cursor {
                    edit(true, cursor, current - cursor);
                }
                cursor = current + len;
            }
            Ok(())
        }
        let mut axes = xlsx_parse::SheetAxes::default();
        restore_axis(&self.rows, MAX_ROWS, |insert, at, count| {
            if insert {
                axes.rows.insert(at, count);
            } else {
                axes.rows.delete(at, count);
            }
        })?;
        restore_axis(&self.cols, MAX_COLS, |insert, at, count| {
            if insert {
                axes.cols.insert(at, count);
            } else {
                axes.cols.delete(at, count);
            }
        })?;
        Ok(axes)
    }
}

#[derive(Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct RecoveryCalculation {
    cells: Vec<Vec<(CellRef, CellState)>>,
    deleted_cells: Vec<Vec<CellRef>>,
    arrays: Vec<Vec<(CellRef, CellRange)>>,
    styles: Option<Stylesheet>,
}

impl RecoveryCalculation {
    fn capture(model: &WorkbookModel, authority: &WorkbookModel) -> Result<Self> {
        if model.sheets.len() != authority.sheets.len() {
            return Err(invalid("Recovery authority sheet count mismatch"));
        }
        let mut cells = Vec::new();
        let mut deleted_cells = Vec::new();
        let mut arrays = Vec::new();
        for (sheet, projected) in model.sheets.iter().zip(&authority.sheets) {
            cells.push(
                sheet
                    .iter_cells()
                    .filter(|&(at, cell)| projected.cell(at) != Some(cell))
                    .map(|(at, cell)| (at, CellState::from(cell)))
                    .collect(),
            );
            deleted_cells.push(
                projected
                    .iter_cells()
                    .filter_map(|(at, _)| sheet.cell(at).is_none().then_some(at))
                    .collect(),
            );
            arrays.push(sheet.array_formulas().collect());
        }
        let state = Self {
            cells,
            deleted_cells,
            arrays,
            styles: (model.styles != authority.styles).then(|| model.styles.clone()),
        };
        if state.apply(authority.clone())? != *model {
            return Err(invalid("Recovery projection cannot be reconstructed"));
        }
        Ok(state)
    }

    fn apply(&self, mut model: WorkbookModel) -> Result<WorkbookModel> {
        let count = model.sheets.len();
        if [
            self.cells.len(),
            self.deleted_cells.len(),
            self.arrays.len(),
        ]
        .iter()
        .any(|length| *length != count)
        {
            return Err(invalid("Recovery calculation sheet count mismatch"));
        }
        if let Some(styles) = &self.styles {
            model.styles = styles.clone();
        }
        for (index, sheet) in model.sheets.iter_mut().enumerate() {
            for (at, cell) in &self.cells[index] {
                validate_cell_ref(*at)?;
                validate_cell_state(cell)?;
                sheet.set_cell(*at, cell.clone().into());
            }
            for at in &self.deleted_cells[index] {
                validate_cell_ref(*at)?;
                sheet.set_cell(*at, xlsx_model::Cell::default());
            }
            let previous: Vec<_> = sheet.array_formulas().map(|(at, _)| at).collect();
            for at in previous {
                sheet.clear_array_formula(at);
            }
            for (at, range) in &self.arrays[index] {
                validate_cell_ref(*at)?;
                validate_range(*range)?;
                sheet.set_array_formula(*at, *range);
            }
        }
        validate_model(&model)?;
        Ok(model)
    }
}

#[derive(Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct RecoveryProposal {
    proposal: Proposal,
    edits: Vec<(String, CellState, Option<NumberFormatMutation>)>,
    ghosts: Vec<(u32, u32, u32, String, String, CellValue)>,
}

impl RecoveryProposal {
    fn capture(proposal: &Proposal) -> Self {
        Self {
            proposal: proposal.clone(),
            edits: proposal
                .edits
                .iter()
                .map(|edit| {
                    (
                        edit.input.clone(),
                        edit.old_state.clone(),
                        edit.number_format.clone(),
                    )
                })
                .collect(),
            ghosts: proposal
                .ghosts
                .iter()
                .map(|ghost| {
                    (
                        ghost.sheet,
                        ghost.row,
                        ghost.col,
                        ghost.old_text.clone(),
                        ghost.new_text.clone(),
                        ghost.alignment_value.clone(),
                    )
                })
                .collect(),
        }
    }

    fn restore(self) -> Result<Proposal> {
        let mut proposal = self.proposal;
        if self.edits.len() != proposal.edits.len() {
            return Err(invalid("Recovery proposal edit count mismatch"));
        }
        for (edit, (input, old_state, number_format)) in proposal.edits.iter_mut().zip(self.edits) {
            edit.input = input;
            edit.old_state = old_state;
            edit.number_format = number_format;
        }
        proposal.ghosts = self
            .ghosts
            .into_iter()
            .map(
                |(sheet, row, col, old_text, new_text, alignment_value)| ProposalGhost {
                    sheet,
                    row,
                    col,
                    old_text,
                    new_text,
                    alignment_value,
                },
            )
            .collect();
        Ok(proposal)
    }
}

fn invalid(message: impl ToString) -> Error {
    Error::InvalidRequest(message.to_string())
}

fn budget() -> SnapshotResult<SnapshotBudget> {
    SnapshotBudget::new(64, crate::snapshot::yrs_split::whole_update_limit() + 12)
}

fn read_sections<'a>(reader: &mut Reader<'a>) -> Result<Vec<&'a [u8]>> {
    let count = reader.var_usize().map_err(invalid)?;
    if count > reader.clone().rest().len() {
        return Err(invalid("Recovery section count exceeds payload"));
    }
    (0..count)
        .map(|_| reader.bytes().map_err(invalid))
        .collect()
}

impl Workbook {
    pub fn recovery_checkpoint(
        &self,
        through_seq: u64,
        calculation_context: Option<CalculationOptions>,
    ) -> Result<Vec<u8>> {
        self.recovery_checkpoint_with_history(through_seq, calculation_context, false)
    }

    pub fn recovery_checkpoint_with_history(
        &self,
        through_seq: u64,
        calculation_context: Option<CalculationOptions>,
        preserve_history: bool,
    ) -> Result<Vec<u8>> {
        if preserve_history && self.is_collaborative() {
            return Err(invalid("History recovery requires a standalone workbook"));
        }
        let budget = budget().map_err(invalid)?;
        let (client_id, _, next_sheet_id) = self.authority.snapshot_identity();
        let header = RecoveryHeader {
            edited_since_open: self.edited_since_open,
            recalculated_since_open: self.recalculated_since_open,
            moved_references_since_open: self.moved_references_since_open,
            active_sheet: self.active_sheet,
            rand_seed: self.rand_seed,
            version_nonce: self.version_nonce.clone(),
            committed_changes: self.committed_changes,
            last_calculation: RecoveryDiagnostics {
                cycle_cells: self.last_calculation.cycle_cells.clone(),
                limited_cells: self.last_calculation.limited_cells.clone(),
            },
            calculation_context,
            client_id,
            next_sheet_id,
        };
        let projection = self.authority.materialize().map_err(authority_error)?;
        let mut preservation = self.preserved.clone();
        preservation.axes = self
            .preserved
            .axes
            .iter()
            .map(|axes| axes.as_ref().map(|_| xlsx_parse::SheetAxes::default()))
            .collect();
        let mut preserved = PreservedSnapshotEncoder::new();
        let mut preserved_chunks = Vec::new();
        while let Some(chunk) = preserved.next(&preservation, budget).map_err(invalid)? {
            preserved_chunks.push(chunk);
        }
        if self.source_package.is_some() && self.source_container.is_none() {
            return Err(invalid("Recovery requires retained package source"));
        }
        let checkpoint = RecoveryCheckpoint {
            engine: env!("CARGO_PKG_VERSION").into(),
            through_seq,
            header,
            calculation: RecoveryCalculation::capture(&self.model, &projection)?,
            collaborative: self.is_collaborative(),
            proposals: self
                .proposals
                .list()
                .iter()
                .map(RecoveryProposal::capture)
                .collect(),
            proposal_id_counter: self.proposals.id_counter(),
            axes: self
                .preserved
                .axes
                .iter()
                .map(|axes| axes.as_ref().map(RecoveryAxes::capture).transpose())
                .collect::<Result<_>>()?,
            history: if preserve_history {
                Some(RecoveryHistory {
                    undo: self.undo.clone(),
                    sheet_order: self.authority.recovery_sheet_order_history(),
                    preserved_undo: self
                        .preserved_undo
                        .iter()
                        .map(RecoveryPreservedHistory::capture)
                        .collect::<Result<_>>()?,
                    preserved_redo: self
                        .preserved_redo
                        .iter()
                        .map(RecoveryPreservedHistory::capture)
                        .collect::<Result<_>>()?,
                })
            } else {
                None
            },
        };
        let mut writer = Writer::new();
        writer.raw(MAGIC);
        writer.u8(CODEC_VERSION);
        writer.raw(&[0; 32]);
        writer.bytes(&serde_json::to_vec(&checkpoint).map_err(invalid)?);
        writer.bytes(&self.authority.encode_recovery_state().map_err(invalid)?);
        writer.var_usize(preserved_chunks.len());
        for chunk in preserved_chunks {
            writer.bytes(&chunk);
        }
        writer.option(self.source_container.as_ref(), |writer, source| {
            writer.bytes(source.as_bytes());
        });
        writer.var_usize(self.pending_remote_updates.len());
        for update in &self.pending_remote_updates {
            writer.bytes(update);
        }
        let mut bytes = writer.into_bytes();
        let digest = Sha256::digest(&bytes[41..]);
        bytes[9..41].copy_from_slice(&digest);
        Ok(bytes)
    }

    pub fn restore_recovery_checkpoint(
        bytes: &[u8],
    ) -> Result<(Self, u64, Option<CalculationOptions>, bool)> {
        if bytes.len() < 41 || &bytes[..8] != MAGIC || bytes[8] != CODEC_VERSION {
            return Err(invalid("Unsupported workbook recovery codec"));
        }
        let payload = &bytes[41..];
        if Sha256::digest(payload)[..] != bytes[9..41] {
            return Err(invalid("Corrupt workbook recovery checkpoint"));
        }
        let mut reader = Reader::new(payload);
        let checkpoint: RecoveryCheckpoint =
            serde_json::from_slice(reader.bytes().map_err(invalid)?).map_err(invalid)?;
        if checkpoint.engine != env!("CARGO_PKG_VERSION") {
            return Err(invalid("Workbook recovery engine version mismatch"));
        }
        let authority_bytes = reader.bytes().map_err(invalid)?;
        let preserved_chunks = read_sections(&mut reader)?;
        let source = reader.option(|reader| reader.bytes()).map_err(invalid)?;
        let pending_remote_updates = read_sections(&mut reader)?;
        reader.finish().map_err(invalid)?;
        let header = checkpoint.header;
        let budget = budget().map_err(invalid)?;
        let mut authority = WorkbookAuthority::restore_recovery_state(
            authority_bytes,
            header.client_id,
            header.next_sheet_id,
        )
        .map_err(invalid)?;
        let projection = authority.materialize().map_err(authority_error)?;
        let model = checkpoint.calculation.apply(projection)?;
        authority.set_snapshot_projection_valid(true);
        let mut preserved = PreservedSnapshotBuilder::new();
        for chunk in preserved_chunks {
            while !preserved
                .advance_bounded(chunk, budget)
                .map_err(invalid)?
                .is_ready()
            {}
        }
        let mut preserved = preserved.finish().map_err(invalid)?;
        if checkpoint.axes.len() != preserved.axes.len()
            || checkpoint
                .axes
                .iter()
                .zip(&preserved.axes)
                .any(|(captured, marker)| captured.is_some() != marker.is_some())
        {
            return Err(invalid("Recovery axis sheet count mismatch"));
        }
        preserved.axes = checkpoint
            .axes
            .into_iter()
            .map(|axes| axes.map(RecoveryAxes::restore).transpose())
            .collect::<Result<_>>()?;
        let count = model.sheets.len();
        if [
            preserved.origins.len(),
            preserved.shared_string_cells.len(),
            preserved.axes.len(),
            preserved.created.len(),
        ]
        .iter()
        .any(|length| *length != count)
        {
            return Err(invalid("Recovery preservation sheet count mismatch"));
        }
        let mut workbook = match source {
            Some(source) => Self::open_for_read(source)?,
            None => Self::from_model(model.clone())?,
        };
        validate_chart_source(&model, workbook.source_package.is_some())?;
        workbook.graph = Some(DepGraph::build(&model));
        workbook.model = model;
        workbook.authority = authority;
        if checkpoint.collaborative {
            workbook.mode = WorkbookMode::Collaborative {
                structure: workbook.authority.structure().map_err(authority_error)?,
            };
        } else if !pending_remote_updates.is_empty() {
            return Err(invalid("Standalone recovery has remote updates"));
        }
        for update in pending_remote_updates {
            validate_collaboration_size(update)?;
            workbook.validate_pending_remote_update(update)?;
            workbook
                .authority
                .stage_updates_v1(&[update], None)
                .map_err(authority_error)?;
            workbook.pending_remote_updates.push(update.to_vec());
        }
        workbook.preserved = preserved;
        workbook.edited_since_open = header.edited_since_open;
        workbook.recalculated_since_open = header.recalculated_since_open;
        workbook.moved_references_since_open = header.moved_references_since_open;
        workbook.set_active_sheet(header.active_sheet)?;
        workbook.rand_seed = header.rand_seed;
        workbook.version_nonce = header.version_nonce;
        workbook.committed_changes = header.committed_changes;
        workbook.last_calculation = CalculationResult {
            changed: Vec::new(),
            cycle_cells: header.last_calculation.cycle_cells,
            limited_cells: header.last_calculation.limited_cells,
        };
        workbook.proposals = ProposalSet::with_id_counter(checkpoint.proposal_id_counter);
        for proposal in checkpoint.proposals {
            workbook.proposals.add(proposal.restore()?);
        }
        let history_reset = checkpoint.history.is_none();
        if let Some(history) = checkpoint.history {
            if checkpoint.collaborative
                || history.undo.undo_depth() != history.preserved_undo.len()
                || history.undo.redo_depth() != history.preserved_redo.len()
            {
                return Err(invalid("Recovery history depth mismatch"));
            }
            workbook
                .authority
                .restore_recovery_sheet_order_history(
                    history.sheet_order,
                    history.undo.undo_depth(),
                    history.undo.redo_depth(),
                )
                .map_err(invalid)?;
            workbook.undo = history.undo;
            workbook.preserved_undo = history
                .preserved_undo
                .into_iter()
                .map(RecoveryPreservedHistory::restore)
                .collect::<Result<_>>()?;
            workbook.preserved_redo = history
                .preserved_redo
                .into_iter()
                .map(RecoveryPreservedHistory::restore)
                .collect::<Result<_>>()?;
        }
        Ok((
            workbook,
            checkpoint.through_seq,
            header.calculation_context,
            history_reset,
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn model() -> WorkbookModel {
        WorkbookModel {
            sheets: vec![xlsx_model::Sheet::new("Data")],
            ..WorkbookModel::default()
        }
    }

    fn structural_round_trip(op: Op) {
        let mut initial = Workbook::from_model(model()).unwrap();
        for (row, col, input) in [(0, 0, "1"), (2, 0, "7"), (0, 2, "11"), (2, 2, "=A3+C1")] {
            initial
                .edit_cell(
                    SheetId(0),
                    CellRef::new(row, col),
                    input,
                    CalculationOptions::default(),
                )
                .unwrap();
        }
        let source = initial.save().unwrap();
        let mut workbook = Workbook::open(&source).unwrap();
        workbook
            .apply_ops(vec![op], CalculationOptions::default())
            .unwrap();
        assert!(!workbook.preserved.axes[0].as_ref().unwrap().is_identity());
        let version = workbook.version();
        let saved = workbook.save().unwrap();
        let bytes = workbook.recovery_checkpoint(1, None).unwrap();
        assert_eq!(workbook.version(), version);
        let (restored, through_seq, _, history_reset) =
            Workbook::restore_recovery_checkpoint(&bytes).unwrap();
        assert_eq!(through_seq, 1);
        assert!(history_reset);
        assert_eq!(restored.version(), version);
        assert_eq!(restored.model, workbook.model);
        assert_eq!(restored.preserved.axes, workbook.preserved.axes);
        assert_eq!(
            restored.sheet_info().unwrap(),
            workbook.sheet_info().unwrap()
        );
        assert_eq!(restored.save().unwrap(), saved);
        assert!(!restored.history_state().can_undo);
        assert!(!restored.history_state().can_redo);
    }

    #[test]
    fn recovery_checkpoint_round_trips_row_insert() {
        structural_round_trip(Op::InsertRows {
            sheet: SheetId(0),
            at: 1,
            count: 2,
        });
    }

    #[test]
    fn recovery_checkpoint_round_trips_row_delete() {
        structural_round_trip(Op::DeleteRows {
            sheet: SheetId(0),
            at: 1,
            count: 2,
        });
    }

    #[test]
    fn recovery_checkpoint_round_trips_column_insert() {
        structural_round_trip(Op::InsertCols {
            sheet: SheetId(0),
            at: 1,
            count: 2,
        });
    }

    #[test]
    fn recovery_checkpoint_round_trips_column_delete() {
        structural_round_trip(Op::DeleteCols {
            sheet: SheetId(0),
            at: 1,
            count: 2,
        });
    }

    #[test]
    fn recovery_axes_round_trip_mixed_edits_and_empty_maps() {
        let mut axes = xlsx_parse::SheetAxes::default();
        axes.rows.insert(3, 2);
        axes.rows.delete(8, 4);
        axes.rows.insert(8, 4);
        axes.rows.delete(3, 2);
        axes.cols.insert(0, 3);
        axes.cols.delete(7, 2);
        assert_eq!(
            RecoveryAxes::capture(&axes).unwrap().restore().unwrap(),
            axes
        );
        axes.rows.delete(0, MAX_ROWS);
        axes.cols.delete(0, MAX_COLS);
        assert_eq!(
            RecoveryAxes::capture(&axes).unwrap().restore().unwrap(),
            axes
        );
    }

    #[test]
    fn recovery_axes_reject_invalid_segments() {
        for rows in [
            vec![(0, 0, 0)],
            vec![(0, 0, MAX_ROWS + 1)],
            vec![(0, 0, 3), (2, 3, 1)],
            vec![(0, 0, 3), (3, 2, 1)],
            vec![(u32::MAX, 0, 1)],
            vec![(0, u32::MAX, 1)],
        ] {
            assert!(
                RecoveryAxes {
                    rows,
                    cols: vec![(0, 0, MAX_COLS)]
                }
                .restore()
                .is_err()
            );
        }
    }

    #[test]
    fn recovery_checkpoint_preserves_calculation_diagnostics_without_notifications() {
        let mut workbook = Workbook::from_model(model()).unwrap();
        workbook
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 0),
                "=A1",
                CalculationOptions::default(),
            )
            .unwrap();
        assert!(!workbook.last_calculation.cycle_cells.is_empty());
        workbook.last_calculation.limited_cells = vec![CellAddress {
            sheet: SheetId(0),
            cell: CellRef::new(1, 0),
        }];
        workbook.last_calculation.changed = (0..1024)
            .map(|row| CellAddress {
                sheet: SheetId(0),
                cell: CellRef::new(row, 0),
            })
            .collect();
        let before = workbook.last_calculation.clone();
        let bytes = workbook.recovery_checkpoint(1, None).unwrap();
        assert_eq!(workbook.last_calculation, before);
        let mut reader = Reader::new(&bytes[41..]);
        let metadata: serde_json::Value = serde_json::from_slice(reader.bytes().unwrap()).unwrap();
        assert!(
            metadata["header"]["last_calculation"]
                .get("changed")
                .is_none()
        );
        let (restored, _, _, _) = Workbook::restore_recovery_checkpoint(&bytes).unwrap();
        assert!(restored.last_calculation.changed.is_empty());
        assert_eq!(restored.last_calculation.cycle_cells, before.cycle_cells);
        assert_eq!(
            restored.last_calculation.limited_cells,
            before.limited_cells
        );
    }

    #[test]
    fn recovery_checkpoint_round_trips_edited_standalone_workbook() {
        let source = Workbook::from_model(model()).unwrap().save().unwrap();
        let mut workbook = Workbook::open(&source).unwrap();
        workbook
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 0),
                "7",
                CalculationOptions::default(),
            )
            .unwrap();
        workbook
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 1),
                "=A1*2",
                CalculationOptions::default(),
            )
            .unwrap();
        workbook
            .patch_range_style(
                SheetId(0),
                CellRange::new(CellRef::new(0, 0), CellRef::new(0, 1)),
                StylePatch {
                    bold: Some(true),
                    ..StylePatch::default()
                },
                CalculationOptions::default(),
            )
            .unwrap();
        assert!(workbook.history_state().can_undo);
        assert!(workbook.last_calculation.changed.is_empty());
        let version = workbook.version();
        let bytes = workbook.recovery_checkpoint(3, None).unwrap();
        assert_eq!(workbook.version(), version);
        let (restored, through_seq, calculation, history_reset) =
            Workbook::restore_recovery_checkpoint(&bytes).unwrap();
        assert_eq!(through_seq, 3);
        assert_eq!(calculation, None);
        assert_eq!(restored.model, workbook.model);
        assert_eq!(restored.last_calculation, workbook.last_calculation);
        assert_eq!(
            restored.sheet_info().unwrap().sheet_ids,
            workbook.sheet_info().unwrap().sheet_ids
        );
        assert_eq!(restored.version(), version);
        assert_eq!(
            restored.encode_state_vector_v1(),
            workbook.encode_state_vector_v1()
        );
        assert_eq!(
            restored.encode_state_as_update_v1(),
            workbook.encode_state_as_update_v1()
        );
        assert_eq!(restored.save().unwrap(), workbook.save().unwrap());
        assert!(!restored.history_state().can_undo);
        assert!(!restored.history_state().can_redo);
        assert_eq!(restored.history_state().undo_depth, 0);
        assert_eq!(restored.history_state().redo_depth, 0);
        assert!(history_reset);
    }

    #[test]
    fn recovery_checkpoint_preserves_undo_redo_and_sheet_identities() {
        let assert_matches = |restored: &Workbook, live: &Workbook| {
            assert_eq!(restored.model, live.model);
            assert_eq!(restored.sheet_info().unwrap(), live.sheet_info().unwrap());
            assert_eq!(
                restored.authority.structure().unwrap().sheet_keys,
                live.authority.structure().unwrap().sheet_keys
            );
            assert_eq!(restored.history_state(), live.history_state());
            assert_eq!(
                serde_json::to_value(&restored.undo).unwrap(),
                serde_json::to_value(&live.undo).unwrap()
            );
        };
        let mut live = Workbook::from_model(model()).unwrap();
        let initial_model = live.model.clone();
        let initial_ids = live.sheet_info().unwrap().sheet_ids;
        live.apply_ops(
            vec![Op::AddSheet {
                index: 0,
                name: "Added".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
        for (sheet, input) in [(SheetId(0), "7"), (SheetId(1), "11")] {
            live.edit_cell(
                sheet,
                CellRef::new(0, 0),
                input,
                CalculationOptions::default(),
            )
            .unwrap();
        }
        let final_model = live.model.clone();
        let final_ids = live.sheet_info().unwrap().sheet_ids;
        let final_undo = serde_json::to_value(&live.undo).unwrap();
        let bytes = live
            .recovery_checkpoint_with_history(64, None, true)
            .unwrap();
        let (mut restored, through_seq, _, history_reset) =
            Workbook::restore_recovery_checkpoint(&bytes).unwrap();
        assert_eq!(through_seq, 64);
        assert!(!history_reset);
        assert_eq!(live.history_state().undo_depth, 3);
        assert_matches(&restored, &live);

        assert!(live.undo(CalculationOptions::default()).unwrap().applied);
        assert!(
            restored
                .undo(CalculationOptions::default())
                .unwrap()
                .applied
        );
        assert_matches(&restored, &live);
        let bytes = live
            .recovery_checkpoint_with_history(65, None, true)
            .unwrap();
        let (mut restored, _, _, history_reset) =
            Workbook::restore_recovery_checkpoint(&bytes).unwrap();
        assert!(!history_reset);
        assert_eq!(restored.history_state().undo_depth, 2);
        assert_eq!(restored.history_state().redo_depth, 1);
        assert_matches(&restored, &live);
        while live.can_undo() {
            assert!(live.undo(CalculationOptions::default()).unwrap().applied);
            assert!(
                restored
                    .undo(CalculationOptions::default())
                    .unwrap()
                    .applied
            );
            assert_matches(&restored, &live);
        }
        assert_eq!(restored.model, initial_model);
        assert_eq!(restored.sheet_info().unwrap().sheet_ids, initial_ids);

        let bytes = live
            .recovery_checkpoint_with_history(67, None, true)
            .unwrap();
        let (mut restored, _, _, history_reset) =
            Workbook::restore_recovery_checkpoint(&bytes).unwrap();
        assert!(!history_reset);
        assert_eq!(restored.history_state().undo_depth, 0);
        assert_eq!(restored.history_state().redo_depth, 3);
        assert_matches(&restored, &live);
        while live.can_redo() {
            assert!(live.redo(CalculationOptions::default()).unwrap().applied);
            assert!(
                restored
                    .redo(CalculationOptions::default())
                    .unwrap()
                    .applied
            );
            assert_matches(&restored, &live);
        }
        assert_eq!(restored.model, final_model);
        assert_eq!(restored.sheet_info().unwrap().sheet_ids, final_ids);
        assert_eq!(serde_json::to_value(&restored.undo).unwrap(), final_undo);
    }

    #[test]
    fn recovery_checkpoint_preserves_sheet_ids_and_proposal_acceptance() {
        let source = Workbook::from_model(model()).unwrap().save().unwrap();
        let mut workbook = Workbook::open(&source).unwrap();
        workbook
            .apply_ops(
                vec![Op::AddSheet {
                    index: 1,
                    name: "Added".into(),
                }],
                CalculationOptions::default(),
            )
            .unwrap();
        let proposal = workbook
            .propose(
                crate::ProposalRequest {
                    agent_id: "agent".into(),
                    note: Some("review".into()),
                    edits: vec![crate::ProposalEditInput {
                        sheet: SheetId(1),
                        cell: CellRef::new(0, 0),
                        input: "=7*2".into(),
                        number_format: None,
                    }],
                },
                CalculationOptions::default(),
            )
            .unwrap();
        let bytes = workbook
            .recovery_checkpoint(
                2,
                Some(CalculationOptions {
                    now_serial: Some(45000.0),
                }),
            )
            .unwrap();
        let (mut restored, _, calculation, _) =
            Workbook::restore_recovery_checkpoint(&bytes).unwrap();
        assert_eq!(
            calculation,
            Some(CalculationOptions {
                now_serial: Some(45000.0)
            })
        );
        assert_eq!(restored.proposals(), workbook.proposals());
        let original_identity = workbook.authority.snapshot_identity();
        let captured_ids = workbook.sheet_info().unwrap().sheet_ids;
        let captured_keys = workbook.authority.structure().unwrap().sheet_keys;
        assert_eq!(captured_ids, vec!["sheet:0", "sheet:1"]);
        assert_eq!(restored.sheet_info().unwrap().sheet_ids, captured_ids);
        assert_eq!(
            restored.authority.structure().unwrap().sheet_keys,
            captured_keys
        );
        let restored_identity = restored.authority.snapshot_identity();
        assert_ne!(restored_identity.0, original_identity.0);
        assert_ne!(restored_identity.1, original_identity.1);
        assert_eq!(restored_identity.2, original_identity.2);
        assert_eq!(restored.save().unwrap(), workbook.save().unwrap());
        restored
            .accept_proposal(&proposal.id, false, CalculationOptions::default())
            .unwrap();
        workbook
            .accept_proposal(&proposal.id, false, CalculationOptions::default())
            .unwrap();
        let ops = vec![Op::AddSheet {
            index: 2,
            name: "Next".into(),
        }];
        restored
            .apply_ops(ops.clone(), CalculationOptions::default())
            .unwrap();
        workbook
            .apply_ops(ops, CalculationOptions::default())
            .unwrap();
        assert_eq!(restored.model, workbook.model);
        let expected_ids = workbook.sheet_info().unwrap().sheet_ids;
        assert_eq!(expected_ids[2], "sheet:2");
        assert_eq!(restored.sheet_info().unwrap().sheet_ids, expected_ids);
        assert!(!captured_ids.contains(&expected_ids[2]));
        assert_eq!(&expected_ids[..2], captured_ids.as_slice());
        let restored_keys = restored.authority.structure().unwrap().sheet_keys;
        assert_eq!(
            restored_keys[2],
            format!("replica:{}:{}", restored_identity.0, restored_identity.2)
        );
        assert!(!captured_keys.contains(&restored_keys[2]));
        assert_eq!(&restored_keys[..2], captured_keys.as_slice());
        assert_eq!(restored.save().unwrap(), workbook.save().unwrap());
    }

    #[test]
    fn recovery_checkpoint_round_trips_collaborative_pending_updates() {
        let mut workbook = Workbook::from_model_collaborative(model(), 71).unwrap();
        let mut peer = Workbook::from_model_collaborative(model(), 72).unwrap();
        let initial = workbook.encode_state_vector_v1();
        peer.edit_cell(
            SheetId(0),
            CellRef::new(0, 0),
            "1",
            CalculationOptions::default(),
        )
        .unwrap();
        let dependency = peer.encode_diff_v1(&initial).unwrap();
        let before = peer.encode_state_vector_v1();
        peer.edit_cell(
            SheetId(0),
            CellRef::new(0, 0),
            "2",
            CalculationOptions::default(),
        )
        .unwrap();
        let pending = peer.encode_diff_v1(&before).unwrap();
        workbook
            .apply_update_v1(&pending, CalculationOptions::default())
            .unwrap();
        assert_eq!(workbook.pending_remote_updates, vec![pending]);
        let bytes = workbook.recovery_checkpoint(5, None).unwrap();
        let (mut restored, through_seq, _, history_reset) =
            Workbook::restore_recovery_checkpoint(&bytes).unwrap();
        assert!(restored.is_collaborative());
        assert_ne!(restored.client_id(), 71);
        let original_identity = workbook.authority.snapshot_identity();
        let restored_identity = restored.authority.snapshot_identity();
        assert_ne!(restored_identity.0, original_identity.0);
        assert_ne!(restored_identity.1, original_identity.1);
        assert_eq!(restored_identity.2, original_identity.2);
        assert_eq!(
            restored.pending_remote_updates,
            workbook.pending_remote_updates
        );
        assert_eq!(
            restored.encode_state_vector_v1(),
            workbook.encode_state_vector_v1()
        );
        assert_eq!(through_seq, 5);
        assert!(history_reset);
        restored
            .apply_update_v1(&dependency, CalculationOptions::default())
            .unwrap();
        workbook
            .apply_update_v1(&dependency, CalculationOptions::default())
            .unwrap();
        assert!(restored.pending_remote_updates.is_empty());
        assert_eq!(restored.model, workbook.model);
        assert_eq!(
            restored.encode_state_vector_v1(),
            peer.encode_state_vector_v1()
        );
    }

    #[test]
    fn recovery_checkpoint_restores_independent_writers_that_converge() {
        let mut workbook = Workbook::from_model_collaborative(model(), 71).unwrap();
        workbook
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 0),
                "0",
                CalculationOptions::default(),
            )
            .unwrap();
        let bytes = workbook.recovery_checkpoint(1, None).unwrap();
        let (mut left, _, _, _) = Workbook::restore_recovery_checkpoint(&bytes).unwrap();
        let (mut right, _, _, _) = Workbook::restore_recovery_checkpoint(&bytes).unwrap();
        assert_ne!(left.client_id(), workbook.client_id());
        assert_ne!(right.client_id(), workbook.client_id());
        assert_ne!(left.client_id(), right.client_id());
        assert_eq!(
            left.encode_state_vector_v1(),
            workbook.encode_state_vector_v1()
        );
        assert_eq!(
            right.encode_state_vector_v1(),
            workbook.encode_state_vector_v1()
        );
        let before = workbook.encode_state_vector_v1();
        left.edit_cell(
            SheetId(0),
            CellRef::new(0, 0),
            "1",
            CalculationOptions::default(),
        )
        .unwrap();
        left.edit_cell(
            SheetId(0),
            CellRef::new(0, 1),
            "11",
            CalculationOptions::default(),
        )
        .unwrap();
        right
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 0),
                "2",
                CalculationOptions::default(),
            )
            .unwrap();
        right
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 2),
                "22",
                CalculationOptions::default(),
            )
            .unwrap();
        let left_update = left.encode_diff_v1(&before).unwrap();
        let right_update = right.encode_diff_v1(&before).unwrap();
        left.apply_update_v1(&right_update, CalculationOptions::default())
            .unwrap();
        right
            .apply_update_v1(&left_update, CalculationOptions::default())
            .unwrap();
        assert_eq!(left.model, right.model);
        assert_eq!(
            left.encode_state_vector_v1(),
            right.encode_state_vector_v1()
        );
        assert_eq!(left.save().unwrap(), right.save().unwrap());
        assert_eq!(
            left.model.sheets[0].cell(CellRef::new(0, 1)).unwrap().value,
            CellValue::Number { value: 11.0 }
        );
        assert_eq!(
            left.model.sheets[0].cell(CellRef::new(0, 2)).unwrap().value,
            CellValue::Number { value: 22.0 }
        );
        let winner = if left.client_id() > right.client_id() {
            1.0
        } else {
            2.0
        };
        assert_eq!(
            left.model.sheets[0].cell(CellRef::new(0, 0)).unwrap().value,
            CellValue::Number { value: winner }
        );
        let (mut forward, _, _, _) = Workbook::restore_recovery_checkpoint(&bytes).unwrap();
        let (mut reverse, _, _, _) = Workbook::restore_recovery_checkpoint(&bytes).unwrap();
        forward
            .apply_update_v1(&left_update, CalculationOptions::default())
            .unwrap();
        forward
            .apply_update_v1(&right_update, CalculationOptions::default())
            .unwrap();
        reverse
            .apply_update_v1(&right_update, CalculationOptions::default())
            .unwrap();
        reverse
            .apply_update_v1(&left_update, CalculationOptions::default())
            .unwrap();
        assert_eq!(forward.model, left.model);
        assert_eq!(reverse.model, left.model);
        assert_eq!(
            forward.encode_state_vector_v1(),
            reverse.encode_state_vector_v1()
        );
    }

    #[test]
    fn recovery_checkpoint_round_trips_chart_workbook() {
        let parts = vec![
            ("xl/workbook.xml".into(), br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_vec()),
            ("xl/_rels/workbook.xml.rels".into(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#.to_vec()),
            ("xl/worksheets/sheet1.xml".into(), br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><drawing r:id="rIdDrawing"/></worksheet>"#.to_vec()),
            ("xl/worksheets/_rels/sheet1.xml.rels".into(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#.to_vec()),
            ("xl/drawings/drawing1.xml".into(), br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>11</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Synthetic"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#.to_vec()),
            ("xl/drawings/_rels/drawing1.xml.rels".into(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#.to_vec()),
            ("xl/charts/chart1.xml".into(), br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Data!$A$1</c:f><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#.to_vec()),
        ];
        let source = ooxml_opc::rezip_parts(&parts).unwrap();
        let mut workbook = Workbook::open(&source).unwrap();
        assert_eq!(workbook.model.sheets[0].charts.len(), 1);
        workbook
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 0),
                "7",
                CalculationOptions::default(),
            )
            .unwrap();
        let bytes = workbook.recovery_checkpoint(1, None).unwrap();
        let (restored, through_seq, _, history_reset) =
            Workbook::restore_recovery_checkpoint(&bytes).unwrap();
        assert_eq!(through_seq, 1);
        assert!(history_reset);
        assert_eq!(restored.model, workbook.model);
        assert_eq!(restored.model.sheets[0].charts.len(), 1);
        assert_eq!(restored.opened_anchors, workbook.opened_anchors);
        assert_eq!(
            restored.encode_state_vector_v1(),
            workbook.encode_state_vector_v1()
        );
        assert_eq!(restored.save().unwrap(), workbook.save().unwrap());
        assert!(!restored.history_state().can_undo);
        assert!(!restored.history_state().can_redo);
    }

    #[test]
    fn recovery_checkpoint_frames_package_and_yrs_bytes_without_json_expansion() {
        let source = Workbook::from_model(model()).unwrap().save().unwrap();
        let mut workbook = Workbook::open(&source).unwrap();
        workbook
            .edit_cell(
                SheetId(0),
                CellRef::new(0, 0),
                "7",
                CalculationOptions::default(),
            )
            .unwrap();
        let bytes = workbook.recovery_checkpoint(1, None).unwrap();
        let mut reader = Reader::new(&bytes[41..]);
        let metadata: serde_json::Value = serde_json::from_slice(reader.bytes().unwrap()).unwrap();
        assert!(metadata.get("source").is_none());
        assert!(metadata.get("authority").is_none());
        assert!(metadata.get("preserved").is_none());
        assert!(metadata.get("pending_remote_updates").is_none());
        assert!(metadata["header"].get("state_vector").is_none());
        assert!(metadata["header"].get("guid").is_none());
        assert_eq!(
            reader.bytes().unwrap(),
            workbook.authority.encode_recovery_state().unwrap()
        );
        read_sections(&mut reader).unwrap();
        assert_eq!(
            reader.option(|reader| reader.bytes()).unwrap(),
            Some(source.as_slice())
        );
        assert!(read_sections(&mut reader).unwrap().is_empty());
        reader.finish().unwrap();
    }

    #[test]
    fn recovery_checkpoint_rejects_wrong_version_or_corrupt_input() {
        let workbook = Workbook::from_model(model()).unwrap();
        let bytes = workbook.recovery_checkpoint(0, None).unwrap();
        let mut wrong_version = bytes.clone();
        wrong_version[8] = 255;
        assert!(Workbook::restore_recovery_checkpoint(&wrong_version).is_err());
        let mut corrupt = bytes.clone();
        let last = corrupt.len() - 1;
        corrupt[last] ^= 1;
        assert!(Workbook::restore_recovery_checkpoint(&corrupt).is_err());
        assert!(Workbook::restore_recovery_checkpoint(&bytes[..bytes.len() - 1]).is_err());
        assert!(Workbook::restore_recovery_checkpoint(&[]).is_err());
        let mut reader = Reader::new(&bytes[41..]);
        let mut checkpoint: RecoveryCheckpoint =
            serde_json::from_slice(reader.bytes().unwrap()).unwrap();
        checkpoint.engine = "unsupported".into();
        let mut writer = Writer::new();
        writer.bytes(&serde_json::to_vec(&checkpoint).unwrap());
        writer.raw(reader.rest());
        let payload = writer.into_bytes();
        let mut wrong_engine = bytes[..9].to_vec();
        wrong_engine.extend_from_slice(&Sha256::digest(&payload));
        wrong_engine.extend_from_slice(&payload);
        assert!(Workbook::restore_recovery_checkpoint(&wrong_engine).is_err());
    }
}
