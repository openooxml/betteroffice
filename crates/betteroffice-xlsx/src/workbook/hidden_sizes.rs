use std::collections::{BTreeMap, BTreeSet, HashMap};

use xlsx_model::{MAX_COLS, MAX_ROWS, Sheet, Workbook as WorkbookModel};
use xlsx_ops::Op;

use crate::authority::base_sheet_index;

type Sizes = BTreeMap<u32, f64>;

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

#[derive(Default)]
pub(crate) struct DimensionChanges {
    rows: BTreeSet<u32>,
    cols: BTreeSet<u32>,
}

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
        axis.extend(at.min(limit)..at.saturating_add(count).min(limit));
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
    source: Option<&WorkbookModel>,
) {
    for (sheet, key) in model.sheets.iter_mut().zip(keys) {
        if let Some(prior) = source
            .and_then(|source| base_sheet_index(key).and_then(|index| source.sheets.get(index)))
        {
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
        } else {
            sheet.hidden_row_heights.clear();
            sheet.hidden_col_widths.clear();
        }
    }
}

pub(super) fn derive_replica_hidden_sizes(
    (model, keys): (&mut WorkbookModel, &[String]),
    (previous, previous_keys): (&mut WorkbookModel, &[String]),
    mut authored: Option<LocalDimensions>,
) {
    let previous_indices = (previous_keys != keys).then(|| {
        previous_keys
            .iter()
            .enumerate()
            .map(|(index, key)| (key.as_str(), index))
            .collect::<HashMap<_, _>>()
    });
    for (index, (sheet, key)) in model.sheets.iter_mut().zip(keys).enumerate() {
        let prior = match &previous_indices {
            Some(indices) => indices.get(key.as_str()).copied(),
            None => Some(index),
        };
        if let Some(prior) = prior {
            let previous = &mut previous.sheets[prior];
            let mut hidden = HiddenSizes::take(previous);
            if authored.is_none() {
                derive_axis(&previous.row_heights, &sheet.row_heights, &mut hidden.rows);
                derive_axis(&previous.col_widths, &sheet.col_widths, &mut hidden.cols);
            }
            hidden.put(sheet);
        }
        if let Some(authored) = authored.as_mut().and_then(|sheets| sheets.remove(&index)) {
            adopt_authored_sizes(sheet, authored);
        }
    }
}

fn derive_axis(previous: &Sizes, installed: &Sizes, hidden: &mut Sizes) {
    if previous == installed {
        return;
    }
    hidden.retain(|at, _| installed.get(at) == Some(&0.0));
    for (&at, &size) in installed {
        if size == 0.0
            && let Some(&size) = previous.get(&at).filter(|&&size| size > 0.0)
        {
            hidden.insert(at, size);
        }
    }
}

pub(super) fn adopt_authored_sizes(sheet: &mut Sheet, authored: AuthoredDimensions) {
    for (hidden, installed, preview, tracks) in [
        (
            &mut sheet.hidden_row_heights,
            &sheet.row_heights,
            &authored.sizes.rows,
            authored.changes.rows,
        ),
        (
            &mut sheet.hidden_col_widths,
            &sheet.col_widths,
            &authored.sizes.cols,
            authored.changes.cols,
        ),
    ] {
        for at in tracks {
            let size = preview
                .get(&at)
                .copied()
                .filter(|_| installed.get(&at) == Some(&0.0));
            if hidden.get(&at).copied() == size {
                continue;
            }
            if let Some(size) = size {
                hidden.insert(at, size);
            } else {
                hidden.remove(&at);
            }
        }
    }
}
