use std::collections::{BTreeMap, BTreeSet, HashMap};

use xlsx_model::{MAX_COLS, MAX_ROWS, Sheet, Workbook as WorkbookModel};
use xlsx_ops::Op;

use crate::authority::base_sheet_index;

type Sizes = BTreeMap<u32, f64>;

#[derive(Default)]
pub(crate) struct HiddenSizes {
    rows: Sizes,
    cols: Sizes,
}

impl HiddenSizes {
    pub(crate) fn take(sheet: &mut Sheet) -> Self {
        Self {
            rows: std::mem::take(&mut sheet.hidden_row_heights),
            cols: std::mem::take(&mut sheet.hidden_col_widths),
        }
    }

    pub(super) fn put(self, sheet: &mut Sheet) {
        sheet.hidden_row_heights = self.rows;
        sheet.hidden_col_widths = self.cols;
    }
}

#[derive(Default)]
pub(crate) struct AxisChanges {
    pub(crate) tracks: BTreeSet<u32>,
    pub(crate) replaced: bool,
}

#[derive(Default)]
pub(crate) struct DimensionChanges {
    pub(crate) rows: AxisChanges,
    pub(crate) cols: AxisChanges,
}

pub(crate) type RemoteDimensions = HashMap<String, DimensionChanges>;

pub(crate) struct AuthoredDimensions {
    pub(crate) sizes: HiddenSizes,
    pub(crate) changes: DimensionChanges,
}

pub(crate) type LocalDimensions = BTreeMap<usize, AuthoredDimensions>;

pub(crate) fn dimension_changes(ops: &[Op]) -> BTreeMap<usize, DimensionChanges> {
    let mut changes = BTreeMap::<_, DimensionChanges>::new();
    for op in ops {
        let (sheet, at, count, rows) = match *op {
            Op::SetRowHeight { sheet, row, .. }
            | Op::SetRowVisibility { sheet, row, .. }
            | Op::RestoreRowDimension { sheet, row, .. } => (sheet, row, 1, true),
            Op::SetRowsVisibility {
                sheet, row, count, ..
            } => (sheet, row, count, true),
            Op::SetColWidth { sheet, col, .. }
            | Op::SetColVisibility { sheet, col, .. }
            | Op::RestoreColDimension { sheet, col, .. } => (sheet, col, 1, false),
            Op::SetColsVisibility {
                sheet, col, count, ..
            } => (sheet, col, count, false),
            _ => continue,
        };
        let sheet = changes.entry(sheet.0 as usize).or_default();
        let (axis, limit) = if rows {
            (&mut sheet.rows, MAX_ROWS)
        } else {
            (&mut sheet.cols, MAX_COLS)
        };
        axis.tracks
            .extend(at.min(limit)..at.saturating_add(count).min(limit));
    }
    changes
}

pub(super) fn with_shared_model<T>(
    model: &mut WorkbookModel,
    read: impl FnOnce(&WorkbookModel) -> T,
) -> T {
    let local = model
        .sheets
        .iter_mut()
        .enumerate()
        .filter(|(_, sheet)| {
            !sheet.hidden_row_heights.is_empty() || !sheet.hidden_col_widths.is_empty()
        })
        .map(|(index, sheet)| (index, HiddenSizes::take(sheet)))
        .collect::<Vec<_>>();
    let result = read(model);
    for (index, sizes) in local {
        sizes.put(&mut model.sheets[index]);
    }
    result
}

pub(super) fn derive_source_hidden_sizes(
    model: &mut WorkbookModel,
    keys: &[String],
    source: &WorkbookModel,
) {
    for (sheet, key) in model.sheets.iter_mut().zip(keys) {
        if let Some(prior) = base_sheet_index(key).and_then(|index| source.sheets.get(index)) {
            let derive = |shared: &Sizes, hidden: &Sizes, visible: &Sizes| {
                shared
                    .iter()
                    .filter(|&(_, &size)| size == 0.0)
                    .filter_map(|(&at, _)| {
                        let size = *hidden.get(&at).or_else(|| visible.get(&at))?;
                        (size.is_finite() && size > 0.0).then_some((at, size))
                    })
                    .collect()
            };
            sheet.hidden_row_heights = derive(
                &sheet.row_heights,
                &prior.hidden_row_heights,
                &prior.row_heights,
            );
            sheet.hidden_col_widths = derive(
                &sheet.col_widths,
                &prior.hidden_col_widths,
                &prior.col_widths,
            );
        }
    }
}

#[derive(Default)]
pub(super) struct ReplicaHiddenSizes {
    remembered: HashMap<String, HiddenSizes>,
    undo: BTreeMap<usize, HiddenSizeStep>,
    redo: BTreeMap<usize, HiddenSizeStep>,
}

pub(super) type HiddenSizeStep = HashMap<String, SheetSizeStep>;

pub(super) struct SheetSizeStep {
    rows: AxisStep,
    cols: AxisStep,
}

type AxisStep = BTreeMap<u32, TrackSizeStep>;

struct TrackSizeStep {
    before: Option<f64>,
    after: Option<f64>,
}

#[derive(Clone, Copy)]
pub(super) enum HiddenSizeHistory<'a> {
    None,
    Record,
    Restore(&'a HiddenSizeStep, bool),
}

pub(super) enum DimensionTransition {
    Local(LocalDimensions),
    Remote(RemoteDimensions),
    Incoming,
}

impl ReplicaHiddenSizes {
    pub(super) fn apply_no_op_sizes(
        &mut self,
        key: &str,
        sheet: &mut Sheet,
        sizes: HiddenSizes,
        mut changes: DimensionChanges,
    ) {
        changes
            .rows
            .tracks
            .retain(|at| sheet.hidden_row_heights.get(at) != sizes.rows.get(at));
        changes
            .cols
            .tracks
            .retain(|at| sheet.hidden_col_widths.get(at) != sizes.cols.get(at));
        if changes.rows.tracks.is_empty() && changes.cols.tracks.is_empty() {
            return;
        }
        self.forget_history(key, &changes);
        if let Some(remembered) = self.remembered.get_mut(key) {
            for at in changes.rows.tracks {
                remembered.rows.remove(&at);
            }
            for at in changes.cols.tracks {
                remembered.cols.remove(&at);
            }
        }
        sizes.put(sheet);
    }

    fn forget_history(&mut self, key: &str, changes: &DimensionChanges) {
        for step in self.undo.values_mut().chain(self.redo.values_mut()) {
            if let Some(sheet) = step.get_mut(key) {
                for (axis, changes) in [
                    (&mut sheet.rows, &changes.rows),
                    (&mut sheet.cols, &changes.cols),
                ] {
                    if changes.replaced {
                        axis.clear();
                    } else {
                        for at in &changes.tracks {
                            axis.remove(at);
                        }
                    }
                }
                if sheet.rows.is_empty() && sheet.cols.is_empty() {
                    step.remove(key);
                }
            }
        }
    }

    pub(super) fn truncate_history(&mut self, undo_depth: usize, redo_depth: usize) {
        for (stack, depth) in [(&mut self.undo, undo_depth), (&mut self.redo, redo_depth)] {
            if stack
                .last_key_value()
                .is_some_and(|(&index, _)| index >= depth)
            {
                drop(stack.split_off(&depth));
            }
        }
    }

    pub(super) fn record(&mut self, depth: usize, step: HiddenSizeStep) {
        if !step.is_empty() {
            self.undo.insert(depth - 1, step);
        }
    }

    pub(super) fn take_step(&mut self, redo: bool, depth: usize) -> HiddenSizeStep {
        let stack = if redo { &mut self.redo } else { &mut self.undo };
        stack.remove(&depth).unwrap_or_default()
    }

    pub(super) fn store_moved_step(&mut self, redo: bool, depth: usize, step: HiddenSizeStep) {
        let stack = if redo { &mut self.undo } else { &mut self.redo };
        if !step.is_empty()
            && let Some(index) = depth.checked_sub(1)
        {
            stack.insert(index, step);
        }
    }

    pub(super) fn reconcile(
        &mut self,
        installed: (&mut WorkbookModel, &[String]),
        previous: (&mut WorkbookModel, &[String]),
        mut transition: DimensionTransition,
        history: HiddenSizeHistory<'_>,
    ) -> HiddenSizeStep {
        if let DimensionTransition::Remote(sheets) = &transition {
            for (key, changes) in sheets {
                self.forget_history(key, changes);
            }
        }
        let (previous, previous_keys) = previous;
        let (model, keys) = installed;
        let previous_indices = (previous_keys != keys).then(|| {
            previous_keys
                .iter()
                .enumerate()
                .map(|(index, key)| (key.as_str(), index))
                .collect::<HashMap<_, _>>()
        });
        let mut step = HiddenSizeStep::new();
        let empty = Sizes::new();
        let empty_step = AxisStep::new();
        for (index, (sheet, key)) in model.sheets.iter_mut().zip(keys).enumerate() {
            let prior = match &previous_indices {
                Some(indices) => indices.get(key.as_str()).copied(),
                None => Some(index),
            };
            let mut prior = prior.map(|index| &mut previous.sheets[index]);
            let mut hidden = prior
                .as_deref_mut()
                .map(HiddenSizes::take)
                .unwrap_or_default();
            let (changes, authored) = match &mut transition {
                DimensionTransition::Local(sheets) => match sheets.remove(&index) {
                    Some(authored) => (Some(authored.changes), Some(authored.sizes)),
                    None => {
                        hidden.put(sheet);
                        continue;
                    }
                },
                DimensionTransition::Remote(sheets) => {
                    let changes = sheets.remove(key);
                    if changes.is_none() && prior.is_some() {
                        hidden.put(sheet);
                        continue;
                    }
                    (changes, None)
                }
                DimensionTransition::Incoming => (None, None),
            };
            let (restore, redo) = match history {
                HiddenSizeHistory::Restore(step, redo) => (step.get(key), redo),
                _ => (None, false),
            };
            let (rows, cols) = prior.map_or((&empty, &empty), |prior| {
                (&prior.row_heights, &prior.col_widths)
            });
            let (row_changes, col_changes) = changes.map_or((None, None), |changes| {
                (Some(changes.rows), Some(changes.cols))
            });
            let mut remembered = self.remembered.remove(key).unwrap_or_default();
            let record = matches!(history, HiddenSizeHistory::Record);
            let row_steps = reconcile_axis(
                (rows, &sheet.row_heights),
                &mut remembered.rows,
                &mut hidden.rows,
                row_changes,
                authored.as_ref().map(|sizes| &sizes.rows),
                (restore.map_or(&empty_step, |step| &step.rows), redo),
                record,
            );
            let col_steps = reconcile_axis(
                (cols, &sheet.col_widths),
                &mut remembered.cols,
                &mut hidden.cols,
                col_changes,
                authored.as_ref().map(|sizes| &sizes.cols),
                (restore.map_or(&empty_step, |step| &step.cols), redo),
                record,
            );
            hidden.put(sheet);
            if !remembered.rows.is_empty() || !remembered.cols.is_empty() {
                self.remembered.insert(key.clone(), remembered);
            }
            if !row_steps.is_empty() || !col_steps.is_empty() {
                step.insert(
                    key.clone(),
                    SheetSizeStep {
                        rows: row_steps,
                        cols: col_steps,
                    },
                );
            }
        }
        if previous_indices.is_some() {
            self.remembered.retain(|key, _| keys.contains(key));
        }
        step
    }
}

fn reconcile_axis(
    (previous, installed): (&Sizes, &Sizes),
    remembered: &mut Sizes,
    hidden: &mut Sizes,
    changes: Option<AxisChanges>,
    authored: Option<&Sizes>,
    (restore, redo): (&AxisStep, bool),
    record: bool,
) -> AxisStep {
    let mut touched = match changes {
        Some(changes) => {
            let mut tracks = changes.tracks;
            if changes.replaced {
                tracks.extend(previous.keys().chain(installed.keys()).copied());
                tracks.extend(hidden.keys().chain(remembered.keys()).copied());
            }
            tracks
        }
        None if previous == installed => BTreeSet::new(),
        None => previous
            .keys()
            .chain(installed.keys())
            .filter(|at| previous.get(*at) != installed.get(*at))
            .copied()
            .collect(),
    };
    touched.extend(restore.keys().copied());
    let mut step = AxisStep::new();
    for at in touched {
        let prior_shared = previous.get(&at).copied();
        let shared = installed.get(&at).copied();
        if let Some(sizes) = authored
            && prior_shared == shared
            && hidden.get(&at) == sizes.get(&at)
        {
            continue;
        }
        let before = hidden.remove(&at).or(remembered.remove(&at));
        let after = if let Some(restored) = restore.get(&at) {
            if redo { restored.after } else { restored.before }
        } else {
            match (authored, shared) {
                (Some(sizes), None | Some(0.0)) => sizes.get(&at).copied(),
                (None, None | Some(0.0)) => before.or(prior_shared.filter(|&size| size > 0.0)),
                _ => None,
            }
        };
        if record && (prior_shared != shared || before != after) {
            step.insert(at, TrackSizeStep { before, after });
        }
        match (shared, after) {
            (Some(0.0), Some(size)) => {
                hidden.insert(at, size);
            }
            (None, Some(size)) => {
                remembered.insert(at, size);
            }
            _ => {}
        }
    }
    step
}
