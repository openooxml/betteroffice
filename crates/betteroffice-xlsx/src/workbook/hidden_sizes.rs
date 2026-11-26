use std::collections::{BTreeMap, HashMap};

use xlsx_model::Workbook as WorkbookModel;

use crate::authority::base_sheet_index;

pub(crate) enum HiddenSizeSource<'a> {
    Base(&'a WorkbookModel),
    Replica(&'a WorkbookModel, &'a [String]),
}

pub(crate) fn derive_hidden_sizes(
    model: &mut WorkbookModel,
    keys: &[String],
    preview: Option<&WorkbookModel>,
    source: Option<HiddenSizeSource<'_>>,
) {
    let keyed_sheets = match &source {
        Some(HiddenSizeSource::Replica(model, keys)) => keys
            .iter()
            .map(String::as_str)
            .zip(&model.sheets)
            .collect::<HashMap<_, _>>(),
        _ => HashMap::new(),
    };
    let from_replica = matches!(&source, Some(HiddenSizeSource::Replica(..)));
    for (index, (sheet, key)) in model.sheets.iter_mut().zip(keys).enumerate() {
        let preview = preview.and_then(|model| model.sheets.get(index));
        let prior = match &source {
            Some(HiddenSizeSource::Base(model)) => {
                base_sheet_index(key).and_then(|index| model.sheets.get(index))
            }
            _ => keyed_sheets.get(key.as_str()).copied(),
        };
        sheet.hidden_col_widths = track_sizes(
            &sheet.col_widths,
            preview.map(|sheet| &sheet.hidden_col_widths),
            prior.map(|sheet| &sheet.hidden_col_widths),
            prior.filter(|_| from_replica).map(|sheet| &sheet.col_widths),
        );
        sheet.hidden_row_heights = track_sizes(
            &sheet.row_heights,
            preview.map(|sheet| &sheet.hidden_row_heights),
            prior.map(|sheet| &sheet.hidden_row_heights),
            prior.filter(|_| from_replica).map(|sheet| &sheet.row_heights),
        );
    }
}

fn track_sizes(
    installed: &BTreeMap<u32, f64>,
    preview: Option<&BTreeMap<u32, f64>>,
    hidden: Option<&BTreeMap<u32, f64>>,
    previous: Option<&BTreeMap<u32, f64>>,
) -> BTreeMap<u32, f64> {
    installed
        .iter()
        .filter(|&(_, &size)| size == 0.0)
        .filter_map(|(&index, _)| {
            let size = preview
                .and_then(|sizes| sizes.get(&index))
                .or_else(|| hidden.and_then(|sizes| sizes.get(&index)))
                .or_else(|| previous.and_then(|sizes| sizes.get(&index)))
                .copied()
                .filter(|size| size.is_finite() && *size > 0.0)?;
            Some((index, size))
        })
        .collect()
}

pub(super) fn equal_ignoring_hidden_sizes(left: &WorkbookModel, right: &mut WorkbookModel) -> bool {
    if left.sheets.len() != right.sheets.len()
        || left.sheets.iter().zip(&right.sheets).any(|(left, right)| {
            left.col_widths != right.col_widths || left.row_heights != right.row_heights
        })
    {
        return false;
    }
    let mut replaced = Vec::new();
    for (index, (left, right)) in left.sheets.iter().zip(&mut right.sheets).enumerate() {
        if left.hidden_col_widths != right.hidden_col_widths
            || left.hidden_row_heights != right.hidden_row_heights
        {
            replaced.push((
                index,
                std::mem::replace(&mut right.hidden_col_widths, left.hidden_col_widths.clone()),
                std::mem::replace(&mut right.hidden_row_heights, left.hidden_row_heights.clone()),
            ));
        }
    }
    let equal = *left == *right;
    for (index, cols, rows) in replaced {
        right.sheets[index].hidden_col_widths = cols;
        right.sheets[index].hidden_row_heights = rows;
    }
    equal
}
