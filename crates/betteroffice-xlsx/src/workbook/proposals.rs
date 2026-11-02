use super::{
    CellRef, Op, Proposal, ProposedEdit, SheetId, Workbook, WorkbookModel, current_cell_state,
    rename_sheet_view,
};

pub(super) fn proposal_edit_is_stale(model: &WorkbookModel, edit: &ProposedEdit) -> bool {
    edit.stale == Some(true)
        || model.sheet(SheetId(edit.sheet)).is_none()
        || current_cell_state(model, SheetId(edit.sheet), CellRef::new(edit.row, edit.col))
            != edit.old_state
}

fn rebases_proposals(op: &Op) -> bool {
    matches!(
        op,
        Op::InsertRows { .. }
            | Op::DeleteRows { .. }
            | Op::InsertCols { .. }
            | Op::DeleteCols { .. }
            | Op::AddSheet { .. }
            | Op::RemoveSheet { .. }
            | Op::MoveSheet { .. }
            | Op::RenameSheet { .. }
            | Op::RestoreSheet { .. }
    )
}

fn remap_address(sheet: &mut u32, row: &mut u32, col: &mut u32, op: &Op) -> bool {
    match *op {
        Op::InsertRows { sheet: target, .. }
        | Op::DeleteRows { sheet: target, .. }
        | Op::InsertCols { sheet: target, .. }
        | Op::DeleteCols { sheet: target, .. }
            if *sheet == target.0 =>
        {
            let Some(at) = xlsx_ops::remap_ref(CellRef::new(*row, *col), op) else {
                return false;
            };
            *row = at.row;
            *col = at.col;
        }
        Op::AddSheet { index, .. } if *sheet >= index as u32 => *sheet += 1,
        Op::RemoveSheet { index } => {
            if *sheet == index as u32 {
                return false;
            }
            if *sheet > index as u32 {
                *sheet -= 1;
            }
        }
        Op::MoveSheet { from, to } => {
            let index = *sheet as usize;
            *sheet = if index == from {
                to
            } else if from < to && index > from && index <= to {
                index - 1
            } else if from > to && index >= to && index < from {
                index + 1
            } else {
                index
            } as u32;
        }
        _ => {}
    }
    true
}

fn rebase_edit(edit: &mut ProposedEdit, formula_input: bool, names: &[String], op: &Op) {
    if edit.stale == Some(true) {
        return;
    }
    let owner = SheetId(edit.sheet);
    if formula_input && let Some(source) = edit.input.strip_prefix('=') {
        match xlsx_ops::remap_proposal_formula(source, owner, names, op) {
            Some(formula) => edit.input = format!("={formula}"),
            None => edit.stale = Some(true),
        }
    }
    if let Some(source) = &edit.old_state.formula {
        match xlsx_ops::remap_proposal_formula(source, owner, names, op) {
            Some(formula) => edit.old_state.formula = Some(formula),
            None => edit.stale = Some(true),
        }
    }
    if !remap_address(&mut edit.sheet, &mut edit.row, &mut edit.col, op) {
        edit.stale = Some(true);
    }
    edit.a1 = CellRef::new(edit.row, edit.col).to_a1();
}

impl Workbook {
    pub(super) fn rebased_proposals(&self, ops: &[Op]) -> Option<Vec<Proposal>> {
        if !ops.iter().any(rebases_proposals) {
            return None;
        }
        let mut proposals = self.proposals.list().to_vec();
        let formula_inputs: Vec<Vec<bool>> = proposals
            .iter()
            .map(|proposal| {
                proposal
                    .edits
                    .iter()
                    .map(|edit| {
                        let style = self.model.sheet(SheetId(edit.sheet)).and_then(|sheet| {
                            sheet.effective_style(CellRef::new(edit.row, edit.col))
                        });
                        let format = self.model.styles.resolved_format(style).number_format;
                        matches!(
                            xlsx_ops::parse_input_with_context(
                                &edit.input,
                                format,
                                self.model.date_system,
                                None,
                            )
                            .0,
                            xlsx_ops::ParsedInput::Formula(source) if xlsx_ops::is_formula_source(&source)
                        )
                    })
                    .collect()
            })
            .collect();
        let mut names = self.sheet_names();
        for op in ops {
            if rebases_proposals(op) {
                for (proposal, formulas) in proposals.iter_mut().zip(&formula_inputs) {
                    for (edit, formula_input) in proposal.edits.iter_mut().zip(formulas) {
                        rebase_edit(edit, *formula_input, &names, op);
                    }
                    proposal.ghosts.retain_mut(|ghost| {
                        remap_address(&mut ghost.sheet, &mut ghost.row, &mut ghost.col, op)
                    });
                }
            }
            rename_sheet_view(&mut names, op);
        }
        for proposal in &mut proposals {
            proposal
                .ghosts
                .sort_by_key(|ghost| (ghost.sheet, ghost.row, ghost.col));
        }
        Some(proposals)
    }

    pub(super) fn adopt_rebased_proposals(&mut self, proposals: Option<Vec<Proposal>>) {
        if let Some(proposals) = proposals {
            for proposal in proposals {
                if let Some(current) = self.proposals.get_mut(&proposal.id) {
                    *current = proposal;
                }
            }
        }
    }

    pub(super) fn restore_rebased_proposals(
        &mut self,
        rebased: Option<Vec<Proposal>>,
        restore: &Option<Vec<Proposal>>,
        reverse: &mut Option<Vec<Proposal>>,
    ) {
        let Some(mut rebased) = rebased else {
            return;
        };
        *reverse = Some(self.proposals.list().to_vec());
        if let Some(restore) = restore {
            for proposal in &mut rebased {
                if let Some(saved) = restore.iter().find(|saved| saved.id == proposal.id) {
                    *proposal = saved.clone();
                }
            }
        }
        self.adopt_rebased_proposals(Some(rebased));
    }
}
