use std::collections::{BTreeSet, HashSet};

use xlsx_calc::graph::DepGraph;
use xlsx_calc::{RecalcResult, rebuild_and_recalc_all, recalc_after};
use xlsx_model::{
    CellRange, CellRef, CellValue, MAX_COLS, MAX_ROWS, Sheet, SheetId, Workbook as WorkbookModel,
};
use xlsx_ops::{
    CellState, Op, Proposal, ProposalSet, ProposedEdit, Provenance, Transaction, UndoStack,
    cell_state_for_input_no_eval,
};
use xlsx_render::{DisplayList, GridGeometry, Viewport, build_display_list, display_text};
#[cfg(feature = "raster")]
use xlsx_render::{scaled, viewport_for_range, viewport_for_used_range};

use crate::authority::{
    AuthorityError, MAX_STATE_VECTOR_ENTRIES, StagedUpdate, SyncOrigin, WorkbookAuthority,
    WorkbookStructure, is_structural_op,
};
use crate::{
    CalculationOptions, CalculationResult, CellAddress, CellEdit, CellInput, Error, MutationResult,
    ProposalAcceptance, ProposalRequest, Result, SheetInfo, UpdateEvent, UpdateOrigin,
};
#[cfg(feature = "raster")]
use crate::{RenderOptions, RenderedPng};

const MAX_RANGE_CELLS: u64 = 100_000;
const MAX_COL_WIDTH: f64 = 255.0;
const MAX_ROW_HEIGHT: f64 = 409.5;
/// Maximum accepted encoded update or state-vector size: 64 MiB.
pub const MAX_COLLABORATION_BYTES: usize = 64 * 1024 * 1024;
/// Largest browser-safe collaboration client identifier.
pub const MAX_COLLABORATION_CLIENT_ID: u64 = (1_u64 << 53) - 1;
/// Maximum client entries accepted in a collaboration state vector.
pub const MAX_COLLABORATION_STATE_VECTOR_ENTRIES: u32 = MAX_STATE_VECTOR_ENTRIES;
const MAX_PENDING_COLLABORATION_UPDATES: usize = 4_096;
pub const MAX_DISPLAY_CELLS: u64 = 250_000;
pub const MAX_PIXMAP_DIM: u32 = 16_384;
pub const MAX_PIXMAP_PIXELS: u64 = 16_777_216;

#[must_use = "dropping the subscription stops update delivery"]
pub struct UpdateSubscription {
    _subscription: yrs::Subscription,
}

enum WorkbookMode {
    Standalone,
    Collaborative { structure: WorkbookStructure },
}

pub struct Workbook {
    authority: WorkbookAuthority,
    mode: WorkbookMode,
    pending_remote_updates: Vec<Vec<u8>>,
    model: WorkbookModel,
    active_sheet: SheetId,
    undo: UndoStack,
    graph: Option<DepGraph>,
    proposals: ProposalSet,
    last_calculation: CalculationResult,
}

impl Workbook {
    pub fn open(bytes: &[u8]) -> Result<Self> {
        Self::open_internal(bytes, true, None)
    }

    pub fn open_for_read(bytes: &[u8]) -> Result<Self> {
        Self::open_internal(bytes, false, None)
    }

    /// Opens a replica. `client_id` must be unique among connected peers.
    pub fn open_collaborative(bytes: &[u8], client_id: u64) -> Result<Self> {
        Self::open_internal(bytes, true, Some(client_id))
    }

    fn open_internal(bytes: &[u8], build_graph: bool, client_id: Option<u64>) -> Result<Self> {
        let parts = ooxml_opc::unzip_parts(bytes).map_err(Error::Package)?;
        let mut names = HashSet::with_capacity(parts.len());
        for (name, _) in &parts {
            if !names.insert(name) {
                return Err(Error::DuplicatePart(name.clone()));
            }
        }
        let model = xlsx_parse::parse_workbook(&parts)?;
        Self::from_parts(model, build_graph, client_id)
    }

    pub fn open_recalculated(bytes: &[u8], options: CalculationOptions) -> Result<Self> {
        let mut workbook = Self::open_internal(bytes, false, None)?;
        workbook.recalculate_all(options);
        Ok(workbook)
    }

    /// Opens and recalculates a replica with a peer-unique client ID.
    pub fn open_collaborative_recalculated(
        bytes: &[u8],
        client_id: u64,
        options: CalculationOptions,
    ) -> Result<Self> {
        let mut workbook = Self::open_internal(bytes, false, Some(client_id))?;
        workbook.recalculate_all(options);
        Ok(workbook)
    }

    pub fn from_model(model: WorkbookModel) -> Result<Self> {
        Self::from_parts(model, true, None)
    }

    /// Creates a replica from a model with a peer-unique client ID.
    pub fn from_model_collaborative(model: WorkbookModel, client_id: u64) -> Result<Self> {
        Self::from_parts(model, true, Some(client_id))
    }

    fn from_parts(model: WorkbookModel, build_graph: bool, client_id: Option<u64>) -> Result<Self> {
        validate_model(&model)?;
        if let Some(client_id) = client_id {
            validate_collaboration_client_id(client_id)?;
        }
        let authority = match client_id {
            Some(client_id) => WorkbookAuthority::from_model_with_client_id(&model, client_id),
            None => WorkbookAuthority::from_model(&model),
        }
        .map_err(authority_error)?;
        let model = authority.materialize().map_err(authority_error)?;
        validate_model(&model)?;
        let graph = build_graph.then(|| DepGraph::build(&model));
        let mode = match client_id {
            Some(_) => WorkbookMode::Collaborative {
                structure: authority.structure().map_err(authority_error)?,
            },
            None => WorkbookMode::Standalone,
        };
        Ok(Self {
            authority,
            mode,
            pending_remote_updates: Vec::new(),
            model,
            active_sheet: SheetId(0),
            undo: UndoStack::new(),
            graph,
            proposals: ProposalSet::new(),
            last_calculation: CalculationResult::default(),
        })
    }

    pub fn client_id(&self) -> u64 {
        self.authority.client_id()
    }

    pub fn is_collaborative(&self) -> bool {
        matches!(self.mode, WorkbookMode::Collaborative { .. })
    }

    pub fn encode_state_vector_v1(&self) -> Vec<u8> {
        self.authority.encode_state_vector_v1()
    }

    pub fn encode_state_as_update_v1(&self) -> Vec<u8> {
        self.authority.encode_state_as_update_v1()
    }

    pub fn encode_diff_v1(&self, remote_state_vector: &[u8]) -> Result<Vec<u8>> {
        validate_collaboration_size(remote_state_vector)?;
        self.authority
            .encode_diff_v1(remote_state_vector)
            .map_err(authority_error)
    }

    pub fn apply_update_v1(
        &mut self,
        update: &[u8],
        options: CalculationOptions,
    ) -> Result<MutationResult> {
        let structure = match &self.mode {
            WorkbookMode::Collaborative { structure } => structure.clone(),
            WorkbookMode::Standalone => return Err(Error::NotCollaborative),
        };
        validate_collaboration_size(update)?;
        if self
            .pending_remote_updates
            .iter()
            .any(|pending| pending == update)
        {
            return Ok(MutationResult::default());
        }
        let pending_bytes = self
            .pending_remote_updates
            .iter()
            .try_fold(update.len(), |total, pending| {
                total.checked_add(pending.len())
            })
            .unwrap_or(usize::MAX);

        if self.pending_remote_updates.is_empty() {
            let staged = self.stage_remote_updates(&[update], &structure)?;
            if staged.pending {
                return self.buffer_pending_remote_update(update);
            }
            return self.apply_staged_remote_update(staged, options, true);
        }

        let mut combined_error = None;
        if pending_bytes <= MAX_COLLABORATION_BYTES {
            let mut updates = self
                .pending_remote_updates
                .iter()
                .map(Vec::as_slice)
                .collect::<Vec<_>>();
            updates.push(update);
            match self.stage_remote_updates(&updates, &structure) {
                Ok(staged) if !staged.pending => {
                    return self.apply_staged_remote_update(staged, options, true);
                }
                Ok(_) => {}
                Err(error) => combined_error = Some(error),
            }
        }

        let staged = self.stage_remote_updates(&[update], &structure)?;
        if staged.pending {
            if let Some(error) = combined_error {
                return Err(error);
            }
            return self.buffer_pending_remote_update(update);
        }
        self.apply_staged_remote_update(staged, options, false)
    }

    fn stage_remote_updates(
        &self,
        updates: &[&[u8]],
        structure: &WorkbookStructure,
    ) -> Result<StagedUpdate> {
        let staged = self
            .authority
            .stage_updates_v1(updates)
            .map_err(authority_error)?;
        if &staged.structure != structure {
            return Err(Error::CollaborativeStructureChanged);
        }
        validate_model(&staged.model)
            .map_err(|error| Error::CollaborativeState(error.to_string()))?;
        Ok(staged)
    }

    fn buffer_pending_remote_update(&mut self, update: &[u8]) -> Result<MutationResult> {
        let updates = self.pending_remote_updates.len() + 1;
        if updates > MAX_PENDING_COLLABORATION_UPDATES {
            return Err(Error::CollaborationPendingUpdatesTooMany {
                updates,
                max: MAX_PENDING_COLLABORATION_UPDATES,
            });
        }
        let bytes = self
            .pending_remote_updates
            .iter()
            .try_fold(update.len(), |total, pending| {
                total.checked_add(pending.len())
            })
            .unwrap_or(usize::MAX);
        if bytes > MAX_COLLABORATION_BYTES {
            return Err(Error::CollaborationDataTooLarge {
                bytes,
                max: MAX_COLLABORATION_BYTES,
            });
        }
        self.pending_remote_updates.push(update.to_vec());
        Ok(MutationResult::default())
    }

    fn apply_staged_remote_update(
        &mut self,
        staged: StagedUpdate,
        options: CalculationOptions,
        resolves_pending: bool,
    ) -> Result<MutationResult> {
        if !staged.effective {
            if resolves_pending {
                self.clear_pending_remote_updates();
            }
            return Ok(MutationResult::default());
        }

        let mut model = staged.model;
        let (graph, recalc) = rebuild_and_recalc_all(&mut model, options.now_serial);
        let mut calculation = calculation_result(&recalc);
        calculation.changed = changed_cells_between(&self.model, &model);
        self.authority
            .apply_update_v1(&staged.update)
            .map_err(authority_error)?;
        if resolves_pending {
            self.clear_pending_remote_updates();
        }
        self.model = model;
        self.graph = Some(graph);
        self.last_calculation = calculation.clone();
        self.undo.clear();
        self.authority.clear_history();
        self.proposals.clear();
        Ok(MutationResult {
            applied: true,
            changed: calculation.changed,
            cycle_cells: calculation.cycle_cells,
            limited_cells: calculation.limited_cells,
        })
    }

    fn clear_pending_remote_updates(&mut self) {
        self.pending_remote_updates.clear();
    }

    pub fn observe_update_v1<F>(&self, callback: F) -> Result<UpdateSubscription>
    where
        F: Fn(UpdateEvent) + 'static,
    {
        let subscription = self
            .authority
            .observe_update_v1(move |remote, update| {
                callback(UpdateEvent {
                    update,
                    origin: if remote {
                        UpdateOrigin::Remote
                    } else {
                        UpdateOrigin::Local
                    },
                });
            })
            .map_err(authority_error)?;
        Ok(UpdateSubscription {
            _subscription: subscription,
        })
    }

    pub fn save(&self) -> Result<Vec<u8>> {
        validate_model(&self.model)?;
        let parts = xlsx_parse::serialize_workbook(&self.model)?;
        ooxml_opc::rezip_parts(&parts).map_err(Error::Package)
    }

    pub fn model(&self) -> &WorkbookModel {
        &self.model
    }

    pub fn into_model(self) -> WorkbookModel {
        self.model
    }

    pub fn sheet(&self, sheet: SheetId) -> Result<&Sheet> {
        self.model.sheet(sheet).ok_or(Error::SheetOutOfRange(sheet))
    }

    pub fn sheet_count(&self) -> usize {
        self.model.sheets.len()
    }

    pub fn sheet_id(&self, name: &str) -> Option<SheetId> {
        self.model.sheet_by_name(name).map(|(id, _)| id)
    }

    pub fn active_sheet(&self) -> SheetId {
        self.active_sheet
    }

    pub fn set_active_sheet(&mut self, sheet: SheetId) -> Result<()> {
        self.sheet(sheet)?;
        self.active_sheet = sheet;
        Ok(())
    }

    pub fn sheet_info(&self) -> Result<SheetInfo> {
        let sheet = self.sheet(self.active_sheet)?;
        let geometry = GridGeometry::new(sheet);
        let (content_width, content_height) = match sheet.used_range() {
            Some(range) => (
                geometry.col_x(range.end.col.saturating_add(2).min(MAX_COLS)),
                geometry.row_y(range.end.row.saturating_add(2).min(MAX_ROWS)),
            ),
            None => (geometry.col_x(26), geometry.row_y(50)),
        };
        Ok(SheetInfo {
            sheet_names: self
                .model
                .sheets
                .iter()
                .map(|sheet| sheet.name.clone())
                .collect(),
            active_sheet: self.active_sheet,
            content_width,
            content_height,
        })
    }

    pub fn cell(&self, sheet: SheetId, cell: CellRef) -> Result<CellEdit> {
        self.validate_cell(cell)?;
        let sheet_ref = self.sheet(sheet)?;
        let (input, is_formula) = match sheet_ref.cell(cell) {
            Some(cell) => match &cell.formula {
                Some(formula) => (format!("={formula}"), true),
                None => (value_to_input(&cell.value), false),
            },
            None => (String::new(), false),
        };
        Ok(CellEdit {
            cell,
            input,
            is_formula,
        })
    }

    pub fn range_cells(&self, sheet: SheetId, range: CellRange) -> Result<Vec<Vec<CellEdit>>> {
        validate_range(range)?;
        self.sheet(sheet)?;
        let rows = u64::from(range.end.row - range.start.row + 1);
        let cols = u64::from(range.end.col - range.start.col + 1);
        if rows * cols > MAX_RANGE_CELLS {
            return Err(Error::RangeTooLarge {
                rows,
                cols,
                max: MAX_RANGE_CELLS,
            });
        }
        let mut cells = Vec::with_capacity(rows as usize);
        for row in range.start.row..=range.end.row {
            let mut row_cells = Vec::with_capacity(cols as usize);
            for col in range.start.col..=range.end.col {
                row_cells.push(self.cell(sheet, CellRef::new(row, col))?);
            }
            cells.push(row_cells);
        }
        Ok(cells)
    }

    pub fn edit_cell(
        &mut self,
        sheet: SheetId,
        cell: CellRef,
        input: &str,
        options: CalculationOptions,
    ) -> Result<MutationResult> {
        self.validate_target(sheet, cell)?;
        let state = edit_cell_state(&self.model, sheet, cell, input);
        validate_cell_state(&state)?;
        if cell_states_semantically_equal(&current_cell_state(&self.model, sheet, cell), &state) {
            return Ok(MutationResult::default());
        }
        self.ensure_graph();
        let formula = state.formula.clone();
        let ops = vec![Op::SetCell {
            sheet,
            at: cell,
            cell: state,
        }];
        self.commit_user(&ops)?;
        self.authority
            .apply_ops(&ops, SyncOrigin::User)
            .map_err(authority_error)?;
        self.graph.as_mut().expect("graph initialized").set_formula(
            sheet,
            cell,
            formula.as_deref(),
        );
        let seeds = [(sheet, cell)];
        let result = recalc_after(
            &mut self.model,
            self.graph.as_mut().expect("graph initialized"),
            &seeds,
            options.now_serial,
        );
        Ok(self.mutation_result(true, result, &seeds))
    }

    pub fn edit_cells(
        &mut self,
        sheet: SheetId,
        edits: &[CellInput],
        options: CalculationOptions,
    ) -> Result<MutationResult> {
        if edits.is_empty() {
            return Ok(MutationResult::default());
        }
        self.sheet(sheet)?;
        let mut touched = Vec::with_capacity(edits.len());
        let mut ops = Vec::with_capacity(edits.len());
        let mut preview = self.model.clone();
        for edit in edits {
            self.validate_cell(edit.cell)?;
            let state = edit_cell_state(&preview, sheet, edit.cell, &edit.input);
            validate_cell_state(&state)?;
            if cell_states_semantically_equal(
                &current_cell_state(&preview, sheet, edit.cell),
                &state,
            ) {
                continue;
            }
            preview
                .sheet_mut(sheet)
                .expect("sheet validated")
                .set_cell(edit.cell, state.clone().into());
            touched.push((sheet, edit.cell, state.formula.clone()));
            ops.push(Op::SetCell {
                sheet,
                at: edit.cell,
                cell: state,
            });
        }
        if ops.is_empty() || models_semantically_equal(&preview, &self.model) {
            return Ok(MutationResult::default());
        }
        self.ensure_graph();
        self.commit_user(&ops)?;
        self.authority
            .apply_ops(&ops, SyncOrigin::User)
            .map_err(authority_error)?;
        for (sheet, cell, formula) in &touched {
            self.graph.as_mut().expect("graph initialized").set_formula(
                *sheet,
                *cell,
                formula.as_deref(),
            );
        }
        let seeds: Vec<_> = touched
            .iter()
            .map(|(sheet, cell, _)| (*sheet, *cell))
            .collect();
        let result = recalc_after(
            &mut self.model,
            self.graph.as_mut().expect("graph initialized"),
            &seeds,
            options.now_serial,
        );
        Ok(self.mutation_result(true, result, &seeds))
    }

    pub fn apply_ops(
        &mut self,
        ops: Vec<Op>,
        options: CalculationOptions,
    ) -> Result<MutationResult> {
        if ops.is_empty() {
            return Ok(MutationResult::default());
        }
        if self.is_collaborative() && ops.iter().any(is_structural_op) {
            return Err(Error::CollaborativeStructureOperation);
        }
        let invalidates_proposals = ops.iter().any(invalidates_proposals);
        let mut preview = self.model.clone();
        for op in &ops {
            validate_op(&preview, op)?;
            validate_insert_capacity(&preview, op)?;
            xlsx_ops::apply(&mut preview, op)?;
            validate_model(&preview)?;
        }
        if preview == self.model {
            return Ok(MutationResult::default());
        }
        let active_name = self.active_sheet_name();
        self.commit_user(&ops)?;
        self.authority
            .apply_ops(&ops, SyncOrigin::User)
            .map_err(authority_error)?;
        self.restore_active_sheet(active_name.as_deref());
        if invalidates_proposals {
            self.proposals.clear();
        }
        let result = self.rebuild_and_recalculate(options);
        Ok(MutationResult {
            applied: true,
            changed: result.changed,
            cycle_cells: result.cycle_cells,
            limited_cells: result.limited_cells,
        })
    }

    pub fn recalculate_all(&mut self, options: CalculationOptions) -> CalculationResult {
        self.rebuild_and_recalculate(options)
    }

    pub fn can_undo(&self) -> bool {
        !self.is_collaborative() && self.undo.can_undo()
    }

    pub fn can_redo(&self) -> bool {
        !self.is_collaborative() && self.undo.can_redo()
    }

    pub fn undo(&mut self, options: CalculationOptions) -> Result<MutationResult> {
        if self.is_collaborative() {
            return Err(Error::CollaborativeUndoUnavailable);
        }
        let active_name = self.active_sheet_name();
        let Some(ops) = self.undo.undo(&mut self.model)? else {
            return Ok(MutationResult::default());
        };
        self.authority
            .apply_ops(&ops, SyncOrigin::Undo)
            .map_err(authority_error)?;
        self.restore_active_sheet(active_name.as_deref());
        if ops.iter().any(invalidates_proposals) {
            self.proposals.clear();
        }
        let result = self.rebuild_and_recalculate(options);
        Ok(MutationResult {
            applied: true,
            changed: result.changed,
            cycle_cells: result.cycle_cells,
            limited_cells: result.limited_cells,
        })
    }

    pub fn redo(&mut self, options: CalculationOptions) -> Result<MutationResult> {
        if self.is_collaborative() {
            return Err(Error::CollaborativeUndoUnavailable);
        }
        let active_name = self.active_sheet_name();
        let Some(ops) = self.undo.redo(&mut self.model)? else {
            return Ok(MutationResult::default());
        };
        self.authority
            .apply_ops(&ops, SyncOrigin::Redo)
            .map_err(authority_error)?;
        self.restore_active_sheet(active_name.as_deref());
        if ops.iter().any(invalidates_proposals) {
            self.proposals.clear();
        }
        let result = self.rebuild_and_recalculate(options);
        Ok(MutationResult {
            applied: true,
            changed: result.changed,
            cycle_cells: result.cycle_cells,
            limited_cells: result.limited_cells,
        })
    }

    pub fn propose(
        &mut self,
        request: ProposalRequest,
        options: CalculationOptions,
    ) -> Result<Proposal> {
        let mut preview = self.model.clone();
        for edit in &request.edits {
            self.validate_target(edit.sheet, edit.cell)?;
            let state = edit_cell_state(&preview, edit.sheet, edit.cell, &edit.input);
            validate_cell_state(&state)?;
            preview
                .sheet_mut(edit.sheet)
                .ok_or(Error::SheetOutOfRange(edit.sheet))?
                .set_cell(edit.cell, state.into());
        }
        rebuild_and_recalc_all(&mut preview, options.now_serial);

        let mut edits = Vec::with_capacity(request.edits.len());
        for edit in request.edits {
            edits.push(ProposedEdit {
                sheet: edit.sheet.0,
                row: edit.cell.row,
                col: edit.cell.col,
                input: edit.input,
                old_state: current_cell_state(&self.model, edit.sheet, edit.cell),
                a1: edit.cell.to_a1(),
                old_text: display_text_at(&self.model, edit.sheet, edit.cell)?,
                new_text: display_text_at(&preview, edit.sheet, edit.cell)?,
            });
        }
        let proposal = Proposal {
            id: self.proposals.next_id(),
            agent_id: request.agent_id,
            note: request.note,
            edits,
        };
        self.proposals.add(proposal.clone());
        Ok(proposal)
    }

    pub fn proposals(&self) -> &[Proposal] {
        self.proposals.list()
    }

    pub fn accept_proposal(
        &mut self,
        id: &str,
        force: bool,
        options: CalculationOptions,
    ) -> Result<ProposalAcceptance> {
        let proposal = self
            .proposals
            .list()
            .iter()
            .find(|proposal| proposal.id == id)
            .cloned()
            .ok_or_else(|| Error::ProposalNotFound(id.to_string()))?;

        if proposal.edits.is_empty() {
            self.proposals.remove(id);
            return Ok(ProposalAcceptance {
                proposal_id: id.to_string(),
                mutation: MutationResult::default(),
            });
        }

        if !force {
            let mut stale = Vec::new();
            for edit in &proposal.edits {
                let address = CellAddress {
                    sheet: SheetId(edit.sheet),
                    cell: CellRef::new(edit.row, edit.col),
                };
                if current_cell_state(&self.model, address.sheet, address.cell) != edit.old_state {
                    stale.push(address);
                }
            }
            if !stale.is_empty() {
                return Err(Error::StaleProposal(stale));
            }
        }

        let mut touched = Vec::with_capacity(proposal.edits.len());
        let mut ops = Vec::with_capacity(proposal.edits.len());
        let mut preview = self.model.clone();
        for edit in &proposal.edits {
            let sheet = SheetId(edit.sheet);
            let cell = CellRef::new(edit.row, edit.col);
            self.validate_target(sheet, cell)?;
            let state = edit_cell_state(&preview, sheet, cell, &edit.input);
            validate_cell_state(&state)?;
            if cell_states_semantically_equal(&current_cell_state(&preview, sheet, cell), &state) {
                continue;
            }
            preview
                .sheet_mut(sheet)
                .expect("sheet validated")
                .set_cell(cell, state.clone().into());
            touched.push((sheet, cell, state.formula.clone()));
            ops.push(Op::SetCell {
                sheet,
                at: cell,
                cell: state,
            });
        }
        if ops.is_empty() || models_semantically_equal(&preview, &self.model) {
            self.proposals.remove(id);
            return Ok(ProposalAcceptance {
                proposal_id: id.to_string(),
                mutation: MutationResult::default(),
            });
        }
        self.ensure_graph();
        self.commit_agent(&ops, proposal.agent_id)?;
        self.authority
            .apply_ops(&ops, SyncOrigin::Agent)
            .map_err(authority_error)?;
        for (sheet, cell, formula) in &touched {
            self.graph.as_mut().expect("graph initialized").set_formula(
                *sheet,
                *cell,
                formula.as_deref(),
            );
        }
        let seeds: Vec<_> = touched
            .iter()
            .map(|(sheet, cell, _)| (*sheet, *cell))
            .collect();
        let result = recalc_after(
            &mut self.model,
            self.graph.as_mut().expect("graph initialized"),
            &seeds,
            options.now_serial,
        );
        let mutation = self.mutation_result(true, result, &seeds);
        self.proposals.remove(id);
        Ok(ProposalAcceptance {
            proposal_id: id.to_string(),
            mutation,
        })
    }

    pub fn reject_proposal(&mut self, id: &str) -> bool {
        self.proposals.remove(id)
    }

    pub fn display_list(&self, viewport: &Viewport) -> Result<DisplayList> {
        self.display_list_for(self.active_sheet, viewport)
    }

    pub fn display_list_for(&self, sheet: SheetId, viewport: &Viewport) -> Result<DisplayList> {
        let sheet_ref = self.sheet(sheet)?;
        validate_display_region(sheet_ref, viewport)?;
        Ok(build_display_list(&self.model, sheet, viewport))
    }

    #[cfg(feature = "raster")]
    pub fn render_png(&self, viewport: &Viewport) -> Result<RenderedPng> {
        self.render_png_for(self.active_sheet, viewport)
    }

    #[cfg(feature = "raster")]
    pub fn render_png_for(&self, sheet: SheetId, viewport: &Viewport) -> Result<RenderedPng> {
        validate_viewport(viewport)?;
        let width = viewport.width.ceil().max(1.0) as u32;
        let height = viewport.height.ceil().max(1.0) as u32;
        validate_render_size(width, height)?;
        let display_list = self.display_list_for(sheet, viewport)?;
        let bytes = xlsx_raster::render_png(&display_list).map_err(Error::Raster)?;
        Ok(RenderedPng {
            bytes,
            width,
            height,
        })
    }

    #[cfg(feature = "raster")]
    pub fn render_sheet(&self, sheet: SheetId, options: &RenderOptions) -> Result<RenderedPng> {
        if !(options.scale.is_finite() && options.scale > 0.0) {
            return Err(Error::InvalidScale(options.scale));
        }
        let sheet_ref = self.sheet(sheet)?;
        if let Some(range) = options.range {
            validate_range(range)?;
        }
        let mut viewport = match options.range {
            Some(range) => viewport_for_range(sheet_ref, range),
            None => viewport_for_used_range(sheet_ref),
        };
        if let Some(width) = options.max_width {
            viewport.width = viewport.width.min(width as f32 / options.scale);
        }
        if let Some(height) = options.max_height {
            viewport.height = viewport.height.min(height as f32 / options.scale);
        }
        validate_viewport(&viewport)?;
        let width = ((viewport.width * options.scale).ceil() as u32).max(1);
        let height = ((viewport.height * options.scale).ceil() as u32).max(1);
        validate_render_size(width, height)?;
        validate_display_region(sheet_ref, &viewport)?;
        let display_list = build_display_list(&self.model, sheet, &viewport);
        let display_list = if options.scale == 1.0 {
            display_list
        } else {
            scaled(display_list, options.scale)
        };
        let bytes = xlsx_raster::render_png(&display_list).map_err(Error::Raster)?;
        Ok(RenderedPng {
            bytes,
            width,
            height,
        })
    }

    pub fn format_address(&self, address: CellAddress) -> String {
        if address.sheet == self.active_sheet {
            address.cell.to_a1()
        } else {
            let name = self
                .model
                .sheet(address.sheet)
                .map(|sheet| sheet.name.as_str())
                .unwrap_or_default();
            format!("{name}!{}", address.cell.to_a1())
        }
    }

    fn validate_target(&self, sheet: SheetId, cell: CellRef) -> Result<()> {
        self.sheet(sheet)?;
        self.validate_cell(cell)
    }

    fn validate_cell(&self, cell: CellRef) -> Result<()> {
        validate_cell_ref(cell)
    }

    fn commit_user(&mut self, ops: &[Op]) -> Result<()> {
        if self.is_collaborative() {
            xlsx_ops::apply_ops(&mut self.model, ops)?;
        } else {
            let transaction = Transaction::new(ops.to_vec(), Provenance::User);
            self.undo.commit(&mut self.model, &transaction)?;
        }
        Ok(())
    }

    fn commit_agent(&mut self, ops: &[Op], agent_id: String) -> Result<()> {
        if self.is_collaborative() {
            xlsx_ops::apply_ops(&mut self.model, ops)?;
        } else {
            let transaction = Transaction::new(ops.to_vec(), Provenance::Agent { id: agent_id });
            self.undo.commit(&mut self.model, &transaction)?;
        }
        Ok(())
    }

    fn rebuild_and_recalculate(&mut self, options: CalculationOptions) -> CalculationResult {
        let (graph, result) = rebuild_and_recalc_all(&mut self.model, options.now_serial);
        self.graph = Some(graph);
        let result = calculation_result(&result);
        self.last_calculation = result.clone();
        result
    }

    fn ensure_graph(&mut self) {
        if self.graph.is_none() {
            self.graph = Some(DepGraph::build(&self.model));
        }
    }

    fn mutation_result(
        &mut self,
        applied: bool,
        result: RecalcResult,
        seeds: &[(SheetId, CellRef)],
    ) -> MutationResult {
        let seeds: HashSet<_> = seeds
            .iter()
            .map(|(sheet, cell)| (sheet.0, cell.row, cell.col))
            .collect();
        self.last_calculation = calculation_result(&result);
        MutationResult {
            applied,
            changed: result
                .changed
                .into_iter()
                .filter(|(sheet, cell)| !seeds.contains(&(sheet.0, cell.row, cell.col)))
                .map(|(sheet, cell)| CellAddress { sheet, cell })
                .collect(),
            cycle_cells: result
                .cycle_cells
                .into_iter()
                .map(|(sheet, cell)| CellAddress { sheet, cell })
                .collect(),
            limited_cells: result
                .limited_cells
                .into_iter()
                .map(|(sheet, cell)| CellAddress { sheet, cell })
                .collect(),
        }
    }

    fn active_sheet_name(&self) -> Option<String> {
        self.model
            .sheet(self.active_sheet)
            .map(|sheet| sheet.name.clone())
    }

    fn restore_active_sheet(&mut self, previous_name: Option<&str>) {
        if let Some(name) = previous_name
            && let Some((sheet, _)) = self.model.sheet_by_name(name)
        {
            self.active_sheet = sheet;
            return;
        }
        let last = self.model.sheets.len().saturating_sub(1) as u32;
        self.active_sheet = SheetId(self.active_sheet.0.min(last));
    }

    pub fn last_calculation(&self) -> &CalculationResult {
        &self.last_calculation
    }
}

fn authority_error(error: AuthorityError) -> Error {
    match error {
        AuthorityError::ClientIdConflict(client_id) => Error::ClientIdConflict(client_id),
        AuthorityError::InvalidStateVector(error) => Error::InvalidStateVector(error),
        AuthorityError::InvalidUpdate(error) => Error::InvalidUpdate(error),
        AuthorityError::InvalidState(error) | AuthorityError::Observer(error) => {
            Error::CollaborativeState(error)
        }
    }
}

fn validate_collaboration_client_id(client_id: u64) -> Result<()> {
    if client_id > MAX_COLLABORATION_CLIENT_ID {
        Err(Error::InvalidClientId {
            client_id,
            max: MAX_COLLABORATION_CLIENT_ID,
        })
    } else {
        Ok(())
    }
}

fn validate_collaboration_size(bytes: &[u8]) -> Result<()> {
    if bytes.len() > MAX_COLLABORATION_BYTES {
        Err(Error::CollaborationDataTooLarge {
            bytes: bytes.len(),
            max: MAX_COLLABORATION_BYTES,
        })
    } else {
        Ok(())
    }
}

fn calculation_result(result: &RecalcResult) -> CalculationResult {
    CalculationResult {
        changed: result
            .changed
            .iter()
            .map(|&(sheet, cell)| CellAddress { sheet, cell })
            .collect(),
        cycle_cells: result
            .cycle_cells
            .iter()
            .map(|&(sheet, cell)| CellAddress { sheet, cell })
            .collect(),
        limited_cells: result
            .limited_cells
            .iter()
            .map(|&(sheet, cell)| CellAddress { sheet, cell })
            .collect(),
    }
}

fn changed_cells_between(before: &WorkbookModel, after: &WorkbookModel) -> Vec<CellAddress> {
    let mut changed = Vec::new();
    for sheet_index in 0..before.sheets.len().max(after.sheets.len()) {
        let sheet = SheetId(sheet_index as u32);
        let mut cells = BTreeSet::new();
        if let Some(before_sheet) = before.sheets.get(sheet_index) {
            cells.extend(
                before_sheet
                    .iter_cells()
                    .map(|(cell, _)| (cell.row, cell.col)),
            );
        }
        if let Some(after_sheet) = after.sheets.get(sheet_index) {
            cells.extend(
                after_sheet
                    .iter_cells()
                    .map(|(cell, _)| (cell.row, cell.col)),
            );
        }
        for (row, col) in cells {
            let cell = CellRef::new(row, col);
            let before_cell = before
                .sheets
                .get(sheet_index)
                .and_then(|sheet| sheet.cell(cell));
            let after_cell = after
                .sheets
                .get(sheet_index)
                .and_then(|sheet| sheet.cell(cell));
            if before_cell != after_cell {
                changed.push(CellAddress { sheet, cell });
            }
        }
    }
    changed
}

fn validate_model(model: &WorkbookModel) -> Result<()> {
    if model.sheets.is_empty() {
        return Err(Error::NoSheets);
    }
    let mut names = HashSet::with_capacity(model.sheets.len());
    for sheet in &model.sheets {
        validate_sheet_name(&sheet.name)?;
        if !names.insert(sheet.name.to_lowercase()) {
            return Err(Error::InvalidOperation(format!(
                "duplicate sheet name: {}",
                sheet.name
            )));
        }
        for (cell, stored) in sheet.iter_cells() {
            validate_cell_ref(cell)?;
            if matches!(stored.value, CellValue::Number { value } if !value.is_finite()) {
                return Err(Error::InvalidOperation(
                    "workbook contains a non-finite cell number".to_string(),
                ));
            }
            if matches!(&stored.value, CellValue::Text { value } if value.chars().count() > xlsx_calc::eval::MAX_CELL_TEXT_CHARS)
            {
                return Err(Error::InvalidOperation(
                    "workbook contains cell text above Excel's length limit".to_string(),
                ));
            }
            if stored
                .formula
                .as_ref()
                .is_some_and(|formula| formula.len() > xlsx_calc::lexer::MAX_FORMULA_BYTES)
            {
                return Err(Error::InvalidOperation(
                    "workbook contains a formula above the length limit".to_string(),
                ));
            }
            if stored
                .style
                .is_some_and(|style| style as usize >= model.styles.cell_xfs.len().max(1))
            {
                return Err(Error::InvalidOperation(
                    "workbook contains an invalid cell style index".to_string(),
                ));
            }
        }
        for (&column, &width) in &sheet.col_widths {
            if column >= MAX_COLS || !width.is_finite() || !(0.0..=MAX_COL_WIDTH).contains(&width) {
                return Err(Error::InvalidOperation(
                    "workbook contains an invalid column width".to_string(),
                ));
            }
        }
        for (&row, &height) in &sheet.row_heights {
            if row >= MAX_ROWS || !height.is_finite() || !(0.0..=MAX_ROW_HEIGHT).contains(&height) {
                return Err(Error::InvalidOperation(
                    "workbook contains an invalid row height".to_string(),
                ));
            }
        }
        for range in &sheet.merges {
            validate_range(*range)?;
        }
    }
    Ok(())
}

fn validate_op(model: &WorkbookModel, op: &Op) -> Result<()> {
    match op {
        Op::SetCell { sheet, at, .. } => {
            require_sheet(model, *sheet)?;
            validate_cell_ref(*at)?;
        }
        Op::InsertRows {
            sheet, at, count, ..
        }
        | Op::DeleteRows {
            sheet, at, count, ..
        } => {
            require_sheet(model, *sheet)?;
            validate_axis("row", *at, *count, MAX_ROWS)?;
        }
        Op::InsertCols {
            sheet, at, count, ..
        }
        | Op::DeleteCols {
            sheet, at, count, ..
        } => {
            require_sheet(model, *sheet)?;
            validate_axis("column", *at, *count, MAX_COLS)?;
        }
        Op::SetColWidth { sheet, col, width } => {
            require_sheet(model, *sheet)?;
            if *col >= MAX_COLS {
                return Err(Error::InvalidOperation(format!(
                    "column {} is out of range",
                    u64::from(*col) + 1
                )));
            }
            if width
                .is_some_and(|width| !width.is_finite() || !(0.0..=MAX_COL_WIDTH).contains(&width))
            {
                return Err(Error::InvalidOperation(format!(
                    "column width must be between 0 and {MAX_COL_WIDTH}"
                )));
            }
        }
        Op::SetRowHeight { sheet, row, height } => {
            require_sheet(model, *sheet)?;
            if *row >= MAX_ROWS {
                return Err(Error::InvalidOperation(format!(
                    "row {} is out of range",
                    u64::from(*row) + 1
                )));
            }
            if height.is_some_and(|height| {
                !height.is_finite() || !(0.0..=MAX_ROW_HEIGHT).contains(&height)
            }) {
                return Err(Error::InvalidOperation(format!(
                    "row height must be between 0 and {MAX_ROW_HEIGHT}"
                )));
            }
        }
        Op::MergeCells { sheet, range } | Op::UnmergeCells { sheet, range } => {
            require_sheet(model, *sheet)?;
            validate_range(*range)?;
        }
        Op::AddSheet { index, .. } => {
            if *index > model.sheets.len() {
                return Err(Error::InvalidOperation(format!(
                    "sheet index {index} out of range"
                )));
            }
        }
        Op::RemoveSheet { index } => {
            if *index >= model.sheets.len() {
                return Err(Error::InvalidOperation(format!(
                    "sheet index {index} out of range"
                )));
            }
        }
        Op::RenameSheet { sheet, .. } => {
            require_sheet(model, *sheet)?;
        }
        Op::RestoreSheet { .. } => {
            return Err(Error::InvalidOperation(
                "restore sheet operations are internal".to_string(),
            ));
        }
    }
    Ok(())
}

fn validate_insert_capacity(model: &WorkbookModel, op: &Op) -> Result<()> {
    match *op {
        Op::InsertRows {
            sheet, at, count, ..
        } => {
            let sheet = require_sheet(model, sheet)?;
            let cutoff = MAX_ROWS - count;
            let loses_cells = sheet
                .iter_cells()
                .any(|(cell, _)| cell.row >= at && cell.row >= cutoff);
            let loses_heights = sheet
                .row_heights
                .keys()
                .any(|&row| row >= at && row >= cutoff);
            let loses_merges = sheet
                .merges
                .iter()
                .any(|range| range.end.row >= at && range.end.row >= cutoff);
            if loses_cells || loses_heights || loses_merges {
                return Err(Error::InvalidOperation(
                    "row insertion would discard content at the sheet boundary".to_string(),
                ));
            }
        }
        Op::InsertCols {
            sheet, at, count, ..
        } => {
            let sheet = require_sheet(model, sheet)?;
            let cutoff = MAX_COLS - count;
            let loses_cells = sheet
                .iter_cells()
                .any(|(cell, _)| cell.col >= at && cell.col >= cutoff);
            let loses_widths = sheet
                .col_widths
                .keys()
                .any(|&col| col >= at && col >= cutoff);
            let loses_merges = sheet
                .merges
                .iter()
                .any(|range| range.end.col >= at && range.end.col >= cutoff);
            if loses_cells || loses_widths || loses_merges {
                return Err(Error::InvalidOperation(
                    "column insertion would discard content at the sheet boundary".to_string(),
                ));
            }
        }
        _ => {}
    }
    Ok(())
}

fn require_sheet(model: &WorkbookModel, sheet: SheetId) -> Result<&Sheet> {
    model.sheet(sheet).ok_or(Error::SheetOutOfRange(sheet))
}

fn validate_sheet_name(name: &str) -> Result<()> {
    let invalid = name.is_empty()
        || name.chars().count() > 31
        || name.starts_with('\'')
        || name.ends_with('\'')
        || name
            .chars()
            .any(|character| matches!(character, ':' | '\\' | '/' | '?' | '*' | '[' | ']'));
    if invalid {
        return Err(Error::InvalidOperation(format!(
            "invalid sheet name: {name:?}"
        )));
    }
    Ok(())
}

fn validate_cell_ref(cell: CellRef) -> Result<()> {
    if cell.row >= MAX_ROWS || cell.col >= MAX_COLS {
        return Err(Error::CellOutOfRange(cell));
    }
    Ok(())
}

fn validate_range(range: CellRange) -> Result<()> {
    validate_cell_ref(range.start)?;
    validate_cell_ref(range.end)?;
    if range.start.row > range.end.row || range.start.col > range.end.col {
        return Err(Error::InvalidOperation(
            "range start must be above and left of range end".to_string(),
        ));
    }
    Ok(())
}

fn validate_cell_state(state: &CellState) -> Result<()> {
    if matches!(state.value, CellValue::Number { value } if !value.is_finite()) {
        return Err(Error::InvalidOperation(
            "cell number must be finite".to_string(),
        ));
    }
    if matches!(&state.value, CellValue::Text { value } if value.chars().count() > xlsx_calc::eval::MAX_CELL_TEXT_CHARS)
    {
        return Err(Error::InvalidOperation(
            "cell text exceeds Excel's length limit".to_string(),
        ));
    }
    if state
        .formula
        .as_ref()
        .is_some_and(|formula| formula.len() > xlsx_calc::lexer::MAX_FORMULA_BYTES)
    {
        return Err(Error::InvalidOperation(
            "formula exceeds the length limit".to_string(),
        ));
    }
    Ok(())
}

fn edit_cell_state(
    workbook: &WorkbookModel,
    sheet: SheetId,
    cell: CellRef,
    input: &str,
) -> CellState {
    let mut state = cell_state_for_input_no_eval(input);
    state.style = workbook
        .sheet(sheet)
        .and_then(|sheet| sheet.cell(cell))
        .and_then(|cell| cell.style);
    state
}

fn current_cell_state(workbook: &WorkbookModel, sheet: SheetId, cell: CellRef) -> CellState {
    workbook
        .sheet(sheet)
        .and_then(|sheet| sheet.cell(cell))
        .map(CellState::from)
        .unwrap_or_default()
}

fn cell_states_semantically_equal(left: &CellState, right: &CellState) -> bool {
    match (&left.formula, &right.formula) {
        (Some(left_formula), Some(right_formula)) => {
            left_formula == right_formula && left.style == right.style
        }
        _ => left == right,
    }
}

fn models_semantically_equal(left: &WorkbookModel, right: &WorkbookModel) -> bool {
    if left.date_system != right.date_system
        || left.shared_strings != right.shared_strings
        || left.styles != right.styles
        || left.sheets.len() != right.sheets.len()
    {
        return false;
    }
    left.sheets.iter().zip(&right.sheets).all(|(left, right)| {
        if left.name != right.name
            || left.merges != right.merges
            || left.col_widths != right.col_widths
            || left.row_heights != right.row_heights
        {
            return false;
        }
        let mut left_cells = left.iter_cells();
        let mut right_cells = right.iter_cells();
        loop {
            match (left_cells.next(), right_cells.next()) {
                (Some((left_at, left_cell)), Some((right_at, right_cell))) => {
                    if left_at != right_at
                        || !cell_states_semantically_equal(
                            &CellState::from(left_cell),
                            &CellState::from(right_cell),
                        )
                    {
                        return false;
                    }
                }
                (None, None) => return true,
                _ => return false,
            }
        }
    })
}

fn display_text_at(workbook: &WorkbookModel, sheet: SheetId, cell: CellRef) -> Result<String> {
    let sheet_ref = workbook.sheet(sheet).ok_or(Error::SheetOutOfRange(sheet))?;
    Ok(match sheet_ref.cell(cell) {
        Some(cell) => display_text(&workbook.styles, workbook.date_system, cell),
        None => String::new(),
    })
}

fn value_to_input(value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Number { value } => value.to_string(),
        CellValue::Bool { value } => if *value { "TRUE" } else { "FALSE" }.to_string(),
        CellValue::Text { value } => {
            if !matches!(xlsx_ops::parse_input(value), xlsx_ops::ParsedInput::Text(text) if text == *value)
            {
                format!("'{value}")
            } else {
                value.clone()
            }
        }
        CellValue::Error { value } => value.as_str().to_string(),
    }
}

fn validate_viewport(viewport: &Viewport) -> Result<()> {
    if !viewport.x.is_finite()
        || !viewport.y.is_finite()
        || !viewport.width.is_finite()
        || !viewport.height.is_finite()
        || viewport.width <= 0.0
        || viewport.height <= 0.0
        || !(viewport.x + viewport.width).is_finite()
        || !(viewport.y + viewport.height).is_finite()
    {
        return Err(Error::InvalidViewport);
    }
    Ok(())
}

fn validate_display_region(sheet: &Sheet, viewport: &Viewport) -> Result<()> {
    validate_viewport(viewport)?;
    let geometry = GridGeometry::new(sheet);
    let right = viewport.x + viewport.width;
    let bottom = viewport.y + viewport.height;
    if right > geometry.col_x(MAX_COLS) || bottom > geometry.row_y(MAX_ROWS) {
        return Err(Error::InvalidViewport);
    }
    let (rows, columns) = geometry.viewport_range(viewport);
    let row_count = u64::from(rows.end - rows.start);
    let column_count = u64::from(columns.end - columns.start);
    let cells = row_count.saturating_mul(column_count);
    if cells > MAX_DISPLAY_CELLS {
        return Err(Error::DisplayTooLarge {
            cells,
            max: MAX_DISPLAY_CELLS,
        });
    }
    Ok(())
}

#[cfg(feature = "raster")]
fn validate_render_size(width: u32, height: u32) -> Result<()> {
    if width > MAX_PIXMAP_DIM || height > MAX_PIXMAP_DIM {
        return Err(Error::RenderTooLarge {
            width,
            height,
            max: MAX_PIXMAP_DIM,
        });
    }
    if u64::from(width) * u64::from(height) > MAX_PIXMAP_PIXELS {
        return Err(Error::RenderAreaTooLarge {
            width,
            height,
            max_pixels: MAX_PIXMAP_PIXELS,
        });
    }
    Ok(())
}

fn validate_axis(axis: &str, at: u32, count: u32, limit: u32) -> Result<()> {
    if count == 0 {
        return Err(Error::InvalidOperation(format!(
            "{axis} operation count must be positive"
        )));
    }
    if at >= limit || count > limit - at {
        return Err(Error::InvalidOperation(format!(
            "{axis} operation exceeds sheet bounds"
        )));
    }
    Ok(())
}

fn invalidates_proposals(op: &Op) -> bool {
    matches!(
        op,
        Op::InsertRows { .. }
            | Op::DeleteRows { .. }
            | Op::InsertCols { .. }
            | Op::DeleteCols { .. }
            | Op::AddSheet { .. }
            | Op::RemoveSheet { .. }
            | Op::RenameSheet { .. }
            | Op::RestoreSheet { .. }
    )
}
