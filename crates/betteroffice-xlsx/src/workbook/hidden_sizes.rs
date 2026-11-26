use std::collections::{BTreeMap, BTreeSet, HashMap};

use xlsx_model::{Sheet, Workbook as WorkbookModel};
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

    fn put(self, sheet: &mut Sheet) {
        sheet.hidden_row_heights = self.rows;
        sheet.hidden_col_widths = self.cols;
    }
}

pub(crate) fn dimension_sheets(ops: &[Op]) -> BTreeSet<usize> {
    ops.iter()
        .filter_map(|op| match op {
            Op::SetRowHeight { sheet, .. }
            | Op::SetRowVisibility { sheet, .. }
            | Op::SetRowsVisibility { sheet, .. }
            | Op::RestoreRowDimension { sheet, .. }
            | Op::SetColWidth { sheet, .. }
            | Op::SetColVisibility { sheet, .. }
            | Op::SetColsVisibility { sheet, .. }
            | Op::RestoreColDimension { sheet, .. } => Some(sheet.0 as usize),
            _ => None,
        })
        .collect()
}

pub(super) fn matches_shared_baseline(
    baseline: &mut WorkbookModel,
    preview: &WorkbookModel,
    sheets: BTreeSet<usize>,
) -> bool {
    let local = sheets
        .into_iter()
        .filter_map(|index| Some((index, HiddenSizes::take(baseline.sheets.get_mut(index)?))))
        .collect::<Vec<_>>();
    let equal = *baseline == *preview;
    for (index, sizes) in local {
        sizes.put(&mut baseline.sheets[index]);
    }
    equal
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
            sheet.hidden_row_heights =
                derive(&sheet.row_heights, &prior.hidden_row_heights, &prior.row_heights);
            sheet.hidden_col_widths =
                derive(&sheet.col_widths, &prior.hidden_col_widths, &prior.col_widths);
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

#[derive(Clone)]
pub(super) struct SheetSizeStep {
    rows: Vec<TrackSizeStep>,
    cols: Vec<TrackSizeStep>,
}

#[derive(Clone, Copy, PartialEq)]
struct TrackSize {
    shared: Option<f64>,
    local: Option<f64>,
}

#[derive(Clone, Copy)]
struct TrackSizeStep {
    at: u32,
    before: TrackSize,
    after: TrackSize,
}

#[derive(Clone, Copy)]
pub(super) enum HiddenSizeHistory<'a> {
    None,
    Record,
    Restore(&'a HiddenSizeStep, bool),
}

impl ReplicaHiddenSizes {
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

    pub(super) fn move_step(
        &mut self,
        redo: bool,
        undo_depth: usize,
        redo_depth: usize,
    ) -> HiddenSizeStep {
        let (from, to, from_depth, to_depth) = if redo {
            (&mut self.redo, &mut self.undo, redo_depth, undo_depth)
        } else {
            (&mut self.undo, &mut self.redo, undo_depth, redo_depth)
        };
        let Some(step) = from.remove(&from_depth) else {
            return HiddenSizeStep::new();
        };
        if let Some(index) = to_depth.checked_sub(1) {
            to.insert(index, step.clone());
        }
        step
    }

    pub(super) fn reconcile(
        &mut self,
        installed: (&mut WorkbookModel, &[String]),
        previous: (&WorkbookModel, &[String]),
        mut authored: BTreeMap<usize, HiddenSizes>,
        history: HiddenSizeHistory<'_>,
    ) -> HiddenSizeStep {
        let previous = previous
            .1
            .iter()
            .map(String::as_str)
            .zip(&previous.0.sheets)
            .collect::<HashMap<_, _>>();
        let mut remembered = std::mem::take(&mut self.remembered);
        let mut step = HiddenSizeStep::new();
        let record = matches!(history, HiddenSizeHistory::Record);
        let (model, keys) = installed;
        for (index, (sheet, key)) in model.sheets.iter_mut().zip(keys).enumerate() {
            let authored = authored.remove(&index);
            let mut memory = remembered.remove(key).unwrap_or_default();
            let (restore, redo) = match history {
                HiddenSizeHistory::Restore(step, redo) => (step.get(key), redo),
                _ => (None, false),
            };
            let prior = previous.get(key.as_str()).copied();
            let empty = Sizes::new();
            let (rows, row_steps) = reconcile_axis(
                &sheet.row_heights,
                prior.map_or((&empty, &empty), |sheet| {
                    (&sheet.row_heights, &sheet.hidden_row_heights)
                }),
                authored.as_ref().map(|sizes| &sizes.rows),
                &mut memory.rows,
                restore.map_or(&[], |step| step.rows.as_slice()),
                redo,
                record,
            );
            let (cols, col_steps) = reconcile_axis(
                &sheet.col_widths,
                prior.map_or((&empty, &empty), |sheet| {
                    (&sheet.col_widths, &sheet.hidden_col_widths)
                }),
                authored.as_ref().map(|sizes| &sizes.cols),
                &mut memory.cols,
                restore.map_or(&[], |step| step.cols.as_slice()),
                redo,
                record,
            );
            (sheet.hidden_row_heights, sheet.hidden_col_widths) = (rows, cols);
            if !memory.rows.is_empty() || !memory.cols.is_empty() {
                self.remembered.insert(key.clone(), memory);
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
        step
    }
}

fn reconcile_axis(
    installed: &Sizes,
    (previous, hidden): (&Sizes, &Sizes),
    authored: Option<&Sizes>,
    remembered: &mut Sizes,
    restore: &[TrackSizeStep],
    redo: bool,
    record: bool,
) -> (Sizes, Vec<TrackSizeStep>) {
    if authored.is_none_or(|authored| authored == hidden)
        && restore.is_empty()
        && installed == previous
    {
        return (hidden.clone(), Vec::new());
    }

    let mut projected = installed
        .iter()
        .filter(|&(_, &size)| size == 0.0)
        .filter_map(|(&at, _)| {
            let size = authored
                .unwrap_or(hidden)
                .get(&at)
                .or_else(|| previous.get(&at).filter(|&&size| size > 0.0))
                .or_else(|| remembered.get(&at))
                .copied()?;
            Some((at, size))
        })
        .collect::<Sizes>();

    let previous_memory = record.then(|| remembered.clone());
    for (&at, &size) in previous {
        if !installed.contains_key(&at)
            && let Some(size) = hidden.get(&at).copied().or((size > 0.0).then_some(size))
        {
            remembered.insert(at, size);
        }
    }
    remembered.retain(|at, _| !installed.contains_key(at));

    for step in restore {
        let (from, to) = if redo {
            (step.before, step.after)
        } else {
            (step.after, step.before)
        };
        if previous.get(&step.at).copied() == from.shared
            && installed.get(&step.at).copied() == to.shared
        {
            let local = match to.shared {
                Some(0.0) => &mut projected,
                None => &mut *remembered,
                _ => continue,
            };
            if let Some(size) = to.local {
                local.insert(step.at, size);
            } else {
                local.remove(&step.at);
            }
        }
    }

    let mut changes = Vec::new();
    if let Some(previous_memory) = previous_memory {
        let tracks = hidden
            .keys()
            .chain(previous_memory.keys())
            .chain(projected.keys())
            .chain(remembered.keys())
            .copied()
            .collect::<BTreeSet<_>>();
        for at in tracks {
            let before = TrackSize {
                shared: previous.get(&at).copied(),
                local: hidden.get(&at).or_else(|| previous_memory.get(&at)).copied(),
            };
            let after = TrackSize {
                shared: installed.get(&at).copied(),
                local: projected.get(&at).or_else(|| remembered.get(&at)).copied(),
            };
            if before != after {
                changes.push(TrackSizeStep { at, before, after });
            }
        }
    }
    (projected, changes)
}
