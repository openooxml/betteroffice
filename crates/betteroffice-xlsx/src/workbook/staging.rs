//! The prepared commit every staged mutation shares: the model change and the authority change
//! are both staged off the live state, and only then is either adopted.

use xlsx_model::Workbook as WorkbookModel;
use xlsx_ops::Op;

use super::{
    PreservedStateHistory, StagedApply, Workbook, WorkbookMode, authority_error,
    retain_array_formulas, retain_formula_caches, validate_collaboration_size,
    validate_collaboration_state,
};
use crate::Result;
use crate::authority::{LocalHistory, StagedLocalUpdate, SyncOrigin};

/// Whether a commit enters undo history.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum CommitHistory {
    Separate,
    None,
}

/// A commit staged against the live state, ready to adopt.
pub(super) struct PreparedCommit {
    pub(super) ops: Vec<Op>,
    /// What the workbook holds once adopted: the staged model standalone, the authority's
    /// projection of it in collaboration.
    pub(super) model: WorkbookModel,
    inverse: Vec<Op>,
    authority: StagedLocalUpdate,
    history: CommitHistory,
}

impl PreparedCommit {
    pub(super) fn update(&self) -> &[u8] {
        &self.authority.update
    }
}

impl Workbook {
    /// Stages `ops`, already applied to `staged`, on a private copy of the authority built from
    /// `baseline` (this replica's encoded state) or a fresh encoding. Nothing live changes.
    pub(super) fn prepare_commit(
        &self,
        ops: Vec<Op>,
        staged: StagedApply,
        origin: SyncOrigin,
        history: CommitHistory,
        baseline: Option<&[u8]>,
    ) -> Result<PreparedCommit> {
        let styles = &staged.model.styles;
        let mut authority = match baseline {
            Some(baseline) => self
                .authority
                .stage_local_ops_from_v1(&ops, origin, styles, baseline),
            None => self.authority.stage_local_ops_v1(&ops, origin, styles),
        }
        .map_err(authority_error)?;
        let model = match &self.mode {
            WorkbookMode::Collaborative { structure } => {
                if &authority.structure != structure {
                    return Err(crate::Error::CollaborativeStructureChanged);
                }
                validate_collaboration_size(&authority.update)?;
                validate_collaboration_state(
                    authority.state_bytes,
                    authority.state_vector_entries,
                )?;
                let mut model = std::mem::take(&mut authority.model);
                retain_formula_caches(&self.model, &mut model);
                retain_array_formulas(&self.model, &mut model);
                model
            }
            WorkbookMode::Standalone => staged.model,
        };
        Ok(PreparedCommit {
            ops,
            model,
            inverse: staged.inverse,
            authority,
            history,
        })
    }

    /// Adopts a prepared commit: the authority first, whose adoption fails before changing
    /// anything, then the model and history, which cannot fail. Returns the update to publish
    /// once the caller has recalculated.
    pub(super) fn commit_prepared(&mut self, prepared: PreparedCommit) -> Result<Option<Vec<u8>>> {
        let PreparedCommit {
            ops,
            model,
            inverse,
            authority,
            history,
        } = prepared;
        let collaborative = self.is_collaborative();
        let recorded = history == CommitHistory::Separate;
        let preserved_before = (!collaborative && recorded).then(|| self.preserved.clone());
        let names_before = self.sheet_names();
        let prior_styles = self.pre_edit_cell_styles(&ops);
        let local_history = match (recorded, collaborative) {
            (true, true) => LocalHistory::Undo,
            (true, false) => LocalHistory::SheetOrder,
            (false, _) => LocalHistory::None,
        };
        let update = self
            .authority
            .adopt_local_update(authority, local_history)
            .map_err(authority_error)?;
        self.install_model(model)?;
        self.update_sheet_info_cache(&ops, &prior_styles);
        if !collaborative && recorded {
            self.undo.record(inverse);
        }
        self.apply_preserved_state_ops(&names_before, &ops);
        if let Some(before) = preserved_before {
            self.preserved_undo.push(PreservedStateHistory {
                before,
                after: self.preserved.clone(),
            });
            self.preserved_redo.clear();
        }
        self.edited_since_open = true;
        Ok(update)
    }
}

#[cfg(test)]
mod tests {
    use std::sync::{Arc, Mutex};

    use xlsx_model::{CellRef, CellValue, Sheet, SheetId, Workbook as WorkbookModel};
    use xlsx_ops::{CellState, Op};

    use super::super::{StagedApply, Workbook};
    use super::CommitHistory;
    use crate::authority::SyncOrigin;
    use crate::{CalculationOptions, ProposalEditInput, ProposalRequest};

    fn model() -> WorkbookModel {
        let mut model = WorkbookModel::default();
        model.sheets.push(Sheet::new("Data"));
        model
    }

    #[test]
    fn a_failing_authority_adoption_changes_nothing() {
        for collaborative in [false, true] {
            let mut workbook = if collaborative {
                Workbook::from_model_collaborative(model(), 51)
            } else {
                Workbook::from_model(model())
            }
            .unwrap();
            let options = CalculationOptions::default();
            workbook
                .edit_cell(SheetId(0), CellRef::new(1, 0), "2", options)
                .unwrap();
            workbook
                .propose(
                    ProposalRequest {
                        agent_id: "agent".into(),
                        note: None,
                        edits: vec![ProposalEditInput {
                            sheet: SheetId(0),
                            cell: CellRef::new(0, 0),
                            input: "7".into(),
                            number_format: None,
                        }],
                    },
                    options,
                )
                .unwrap();
            let events = Arc::new(Mutex::new(0));
            let counted = Arc::clone(&events);
            let _subscription = workbook
                .observe_update_v1(move |_| *counted.lock().unwrap() += 1)
                .unwrap();
            let op = Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::new(0, 0),
                cell: CellState {
                    value: CellValue::Number { value: 9.0 },
                    ..CellState::default()
                },
            };
            let mut preview = workbook.model().clone();
            let inverse = xlsx_ops::apply_in_place(&mut preview, &op).unwrap().0;
            let mut prepared = workbook
                .prepare_commit(
                    vec![op],
                    StagedApply::new(preview, inverse),
                    SyncOrigin::Agent,
                    CommitHistory::Separate,
                    None,
                )
                .unwrap();
            prepared.authority.update = vec![0xff, 0xff, 0xff];
            let model = workbook.model().clone();
            let state = workbook.encode_state_as_update_v1();
            let history = workbook.history_state();
            let proposals = workbook.proposals().to_vec();
            let version = workbook.version();

            assert!(workbook.commit_prepared(prepared).is_err());
            assert_eq!(workbook.model(), &model);
            assert_eq!(workbook.encode_state_as_update_v1(), state);
            assert_eq!(workbook.history_state(), history);
            assert_eq!(workbook.proposals(), proposals.as_slice());
            assert_eq!(workbook.version(), version);
            assert_eq!(*events.lock().unwrap(), 0);
            workbook.undo(options).unwrap();
            assert_eq!(
                workbook.cell(SheetId(0), CellRef::new(1, 0)).unwrap().input,
                ""
            );
        }
    }
}
