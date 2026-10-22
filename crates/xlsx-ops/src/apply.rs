//! applying ops mutably to a workbook, returning the inverse for undo, plus
//! `remap_ref` — the shared address-remapping primitive.

use std::collections::BTreeMap;
use std::fmt;

use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::workbook::MAX_DEFINED_NAMES;
use xlsx_model::{
    Cell, CellRange, CellRef, ColId, ConditionalFormatRule, ConditionalFormatting, DefinedName,
    DifferentialFormat, RowId, Sheet, SheetId, Workbook,
};

use crate::formatting::{mutate_number_format, patch_cell_format};
use crate::op::{CellState, Op};
use crate::remap::{
    remap_auto_filter, remap_charts, remap_defined_names, remap_formulas,
    remap_hyperlink_locations, remap_hyperlink_range, rename_chart_refs, rename_defined_names,
    rename_hyperlink_location, rename_sheet_references, strand_chart_refs,
};

/// the inverse of an applied op: a base-vocabulary op list that, replayed in
/// order, restores the prior workbook state.
#[derive(Debug, Clone, PartialEq)]
pub struct InvertedOp(pub Vec<Op>);

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OpError {
    SheetNotFound(SheetId),
    SheetIndexOutOfRange(usize),
    FormulaNotRewritable { sheet: SheetId, cell: CellRef },
    DefinedNameNotRewritable { name: String },
    ChartRefNotRewritable { part: String },
    ChartAnchorNotMovable { part: String },
    ChartNotFound { frame: String },
    ChartFrameShifted { frame: String },
    NumFmtTableFull,
    InvalidStyle(String),
    InvalidConditionalFormat(String),
    InvalidAutoFilter(String),
    ConditionalFormatNotFound(String),
    InvalidDefinedName(String),
    TooManyDefinedNames,
}

impl fmt::Display for OpError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            OpError::SheetNotFound(id) => write!(f, "sheet {} not found", id.0),
            OpError::NumFmtTableFull => write!(f, "number format table is full"),
            OpError::TooManyDefinedNames => write!(f, "defined name count exceeded cap"),
            OpError::SheetIndexOutOfRange(i) => write!(f, "sheet index {i} out of range"),
            OpError::FormulaNotRewritable { sheet, cell } => write!(
                f,
                "formula at sheet {}, {} cannot be safely rewritten",
                sheet.0,
                cell.to_a1()
            ),
            OpError::DefinedNameNotRewritable { name } => {
                write!(f, "defined name {name} cannot be safely rewritten")
            }
            OpError::ChartRefNotRewritable { part } => {
                write!(f, "chart reference in {part} cannot be safely rewritten")
            }
            OpError::ChartAnchorNotMovable { part } => {
                write!(f, "the anchor of {part} would leave the sheet grid")
            }
            OpError::ChartNotFound { frame } => {
                write!(f, "no chart frame on this sheet is named {frame}")
            }
            OpError::ChartFrameShifted { frame } => write!(
                f,
                "chart frame {frame} does not hold the chart part and anchor this op was recorded against; its drawing has changed"
            ),
            OpError::InvalidStyle(message) | OpError::InvalidDefinedName(message) => {
                f.write_str(message)
            }
            OpError::InvalidAutoFilter(message) | OpError::InvalidConditionalFormat(message) => {
                f.write_str(message)
            }
            OpError::ConditionalFormatNotFound(id) => {
                write!(f, "conditional format {id} not found")
            }
        }
    }
}

impl std::error::Error for OpError {}

/// apply one op, returning its inverse; refused structural ops leave `wb`
/// untouched.
pub fn apply(wb: &mut Workbook, op: &Op) -> Result<InvertedOp, OpError> {
    match op {
        Op::InsertRows { .. }
        | Op::DeleteRows { .. }
        | Op::InsertCols { .. }
        | Op::DeleteCols { .. } => apply_atomically(wb, |next| apply_in_place(next, op)),
        _ => apply_in_place(wb, op),
    }
}

/// apply one op directly with no rollback clone; only for models the caller
/// drops on error.
pub fn apply_in_place(wb: &mut Workbook, op: &Op) -> Result<InvertedOp, OpError> {
    crate::conditional_formats::ensure_conditional_formats_allow_edit(wb, op)?;
    match op {
        Op::SetCell { sheet, at, cell } => {
            wb.styles.preserve_style_identity();
            let s = sheet_mut(wb, *sheet)?;
            let old = s.cell(*at).map(CellState::from).unwrap_or_default();
            s.set_cell(*at, cell.clone().into());
            Ok(InvertedOp(vec![Op::SetCell {
                sheet: *sheet,
                at: *at,
                cell: old,
            }]))
        }
        Op::SetColWidth { sheet, col, width } => {
            let s = sheet_mut(wb, *sheet)?;
            let old = s.col_widths.get(col).copied();
            let old_hidden = s.hidden_col_widths.get(col).copied();
            match width {
                Some(w) => {
                    if *w == 0.0 {
                        if let Some(previous) = old.filter(|previous| *previous > 0.0) {
                            s.hidden_col_widths.insert(*col, previous);
                        }
                    } else {
                        s.hidden_col_widths.remove(col);
                    }
                    s.col_widths.insert(*col, *w);
                }
                None => {
                    s.col_widths.remove(col);
                    s.hidden_col_widths.remove(col);
                }
            }
            Ok(InvertedOp(vec![Op::RestoreColDimension {
                sheet: *sheet,
                col: *col,
                width: old,
                hidden_width: old_hidden,
            }]))
        }
        Op::SetColVisibility { sheet, col, hidden } => {
            let s = sheet_mut(wb, *sheet)?;
            let old = s.col_widths.get(col).copied();
            let old_hidden = s.hidden_col_widths.get(col).copied();
            if *hidden {
                if let Some(size) = old.filter(|size| *size > 0.0) {
                    s.hidden_col_widths.insert(*col, size);
                }
                s.col_widths.insert(*col, 0.0);
            } else if old == Some(0.0) {
                if let Some(size) = s.hidden_col_widths.remove(col) {
                    s.col_widths.insert(*col, size);
                } else {
                    s.col_widths.remove(col);
                }
            }
            Ok(InvertedOp(vec![Op::RestoreColDimension {
                sheet: *sheet,
                col: *col,
                width: old,
                hidden_width: old_hidden,
            }]))
        }
        Op::RestoreColDimension {
            sheet,
            col,
            width,
            hidden_width,
        } => {
            let s = sheet_mut(wb, *sheet)?;
            let old = s.col_widths.get(col).copied();
            let old_hidden = s.hidden_col_widths.get(col).copied();
            if let Some(width) = width {
                s.col_widths.insert(*col, *width);
            } else {
                s.col_widths.remove(col);
            }
            if let Some(width) = hidden_width {
                s.hidden_col_widths.insert(*col, *width);
            } else {
                s.hidden_col_widths.remove(col);
            }
            Ok(InvertedOp(vec![Op::RestoreColDimension {
                sheet: *sheet,
                col: *col,
                width: old,
                hidden_width: old_hidden,
            }]))
        }
        Op::SetRowHeight { sheet, row, height } => {
            let s = sheet_mut(wb, *sheet)?;
            let old = s.row_heights.get(row).copied();
            let old_hidden = s.hidden_row_heights.get(row).copied();
            match height {
                Some(h) => {
                    if *h == 0.0 {
                        if let Some(previous) = old.filter(|previous| *previous > 0.0) {
                            s.hidden_row_heights.insert(*row, previous);
                        }
                    } else {
                        s.hidden_row_heights.remove(row);
                    }
                    s.row_heights.insert(*row, *h);
                }
                None => {
                    s.row_heights.remove(row);
                    s.hidden_row_heights.remove(row);
                }
            }
            Ok(InvertedOp(vec![Op::RestoreRowDimension {
                sheet: *sheet,
                row: *row,
                height: old,
                hidden_height: old_hidden,
            }]))
        }
        Op::SetRowVisibility { sheet, row, hidden } => {
            let s = sheet_mut(wb, *sheet)?;
            let old = s.row_heights.get(row).copied();
            let old_hidden = s.hidden_row_heights.get(row).copied();
            if *hidden {
                if let Some(size) = old.filter(|size| *size > 0.0) {
                    s.hidden_row_heights.insert(*row, size);
                }
                s.row_heights.insert(*row, 0.0);
            } else if s.row_hidden(*row) {
                if let Some(size) = s.hidden_row_heights.remove(row) {
                    s.row_heights.insert(*row, size);
                } else {
                    s.row_heights.remove(row);
                    if s.row_hidden(*row) {
                        let height = s
                            .format
                            .default_row_height_pt
                            .filter(|height| *height > 0.0)
                            .unwrap_or(15.0);
                        s.row_heights.insert(*row, height);
                    }
                }
            }
            Ok(InvertedOp(vec![Op::RestoreRowDimension {
                sheet: *sheet,
                row: *row,
                height: old,
                hidden_height: old_hidden,
            }]))
        }
        Op::RestoreRowDimension {
            sheet,
            row,
            height,
            hidden_height,
        } => {
            let s = sheet_mut(wb, *sheet)?;
            let old = s.row_heights.get(row).copied();
            let old_hidden = s.hidden_row_heights.get(row).copied();
            if let Some(height) = height {
                s.row_heights.insert(*row, *height);
            } else {
                s.row_heights.remove(row);
            }
            if let Some(height) = hidden_height {
                s.hidden_row_heights.insert(*row, *height);
            } else {
                s.hidden_row_heights.remove(row);
            }
            Ok(InvertedOp(vec![Op::RestoreRowDimension {
                sheet: *sheet,
                row: *row,
                height: old,
                hidden_height: old_hidden,
            }]))
        }
        Op::SetFreezePane { sheet, pane } => {
            let sheet_ref = sheet_mut(wb, *sheet)?;
            let old = sheet_ref.freeze_pane;
            sheet_ref.freeze_pane = *pane;
            Ok(InvertedOp(vec![Op::SetFreezePane {
                sheet: *sheet,
                pane: old,
            }]))
        }
        Op::SetHyperlinks { sheet, hyperlinks } => {
            let sheet_ref = sheet_mut(wb, *sheet)?;
            let old = std::mem::replace(&mut sheet_ref.hyperlinks, hyperlinks.clone());
            Ok(InvertedOp(vec![Op::SetHyperlinks {
                sheet: *sheet,
                hyperlinks: old,
            }]))
        }
        Op::SetHyperlink { sheet, hyperlink } => {
            let mut hyperlink = hyperlink.clone();
            hyperlink.range = CellRange {
                start: CellRef::new(hyperlink.range.start.row, hyperlink.range.start.col),
                end: CellRef::new(hyperlink.range.end.row, hyperlink.range.end.col),
            };
            let links = &mut sheet_mut(wb, *sheet)?.hyperlinks;
            let replaced = links
                .iter()
                .filter(|existing| existing.range.overlaps(&hyperlink.range))
                .collect::<Vec<_>>();
            if replaced.as_slice() == [&hyperlink] {
                return Ok(InvertedOp(vec![]));
            }
            let exact = replaced.len() == 1 && replaced[0].range == hyperlink.range;
            let old = links.clone();
            if exact {
                let index = links
                    .iter()
                    .position(|link| link.range == hyperlink.range)
                    .unwrap();
                links[index] = hyperlink;
            } else {
                links.retain(|existing| !existing.range.overlaps(&hyperlink.range));
                links.push(hyperlink);
            }
            Ok(InvertedOp(vec![Op::SetHyperlinks {
                sheet: *sheet,
                hyperlinks: old,
            }]))
        }
        Op::RemoveHyperlink { sheet, range } => {
            let range = CellRange {
                start: CellRef::new(range.start.row, range.start.col),
                end: CellRef::new(range.end.row, range.end.col),
            };
            let links = &mut sheet_mut(wb, *sheet)?.hyperlinks;
            if !links.iter().any(|link| link.range == range) {
                return Ok(InvertedOp(vec![]));
            }
            let old = links.clone();
            links.retain(|link| link.range != range);
            Ok(InvertedOp(vec![Op::SetHyperlinks {
                sheet: *sheet,
                hyperlinks: old,
            }]))
        }
        Op::AddConditionalFormat {
            sheet,
            range,
            kind,
            operator,
            formulas,
            format,
        } => {
            let formulas =
                validate_conditional_format(kind, operator.as_deref(), formulas, format)?;
            let s = sheet_mut(wb, *sheet)?;
            if range.end.row >= MAX_ROWS || range.end.col >= MAX_COLS {
                return Err(OpError::InvalidConditionalFormat(
                    "conditional format range is outside the sheet grid".into(),
                ));
            }
            let old = s.conditional_formats.clone();
            let priority = old
                .iter()
                .flat_map(|group| &group.rules)
                .map(|rule| rule.priority)
                .max()
                .unwrap_or(0)
                .checked_add(1)
                .ok_or_else(|| {
                    OpError::InvalidConditionalFormat("conditional format priority overflow".into())
                })?;
            let rule = ConditionalFormatRule {
                id: next_conditional_format_id(&old),
                rule_type: kind.clone(),
                operator: operator.clone(),
                formulas,
                priority,
                dxf_id: None,
                format: format.clone(),
                raw_xml: None,
            };
            let sqref = range.to_a1();
            match s
                .conditional_formats
                .iter_mut()
                .find(|group| group.sqref == sqref)
            {
                Some(group) => group.rules.push(rule),
                None => s.conditional_formats.push(ConditionalFormatting {
                    sqref,
                    rules: vec![rule],
                }),
            }
            Ok(InvertedOp(vec![Op::SetConditionalFormats {
                sheet: *sheet,
                formats: old,
            }]))
        }
        Op::RemoveConditionalFormat { sheet, id } => {
            let s = sheet_mut(wb, *sheet)?;
            let found = s
                .conditional_formats
                .iter()
                .flat_map(|group| &group.rules)
                .find(|rule| &rule.id == id);
            match found {
                None => return Err(OpError::ConditionalFormatNotFound(id.clone())),
                Some(rule) if !rule.editable() => {
                    return Err(OpError::InvalidConditionalFormat(format!(
                        "conditional format {id} of type {} is not editable",
                        rule.rule_type
                    )));
                }
                _ => {}
            }
            let old = s.conditional_formats.clone();
            for group in &mut s.conditional_formats {
                group.rules.retain(|rule| &rule.id != id);
            }
            s.conditional_formats
                .retain(|group| !group.rules.is_empty());
            Ok(InvertedOp(vec![Op::SetConditionalFormats {
                sheet: *sheet,
                formats: old,
            }]))
        }
        Op::SetAutoFilter { sheet, range } => {
            let current = wb.sheet(*sheet).ok_or(OpError::SheetNotFound(*sheet))?;
            let previous = current.auto_filter.clone();
            if previous
                .as_ref()
                .is_some_and(|filter| filter.source_count > 1)
            {
                return Err(OpError::InvalidAutoFilter("This sheet has multiple autoFilter elements; repair the source before setting its filter.".into()));
            }
            if previous.is_none() && range.is_none() {
                return Ok(InvertedOp(Vec::new()));
            }
            if let Some(range) = range {
                if range.start.row > range.end.row
                    || range.start.col > range.end.col
                    || range.end.row >= MAX_ROWS
                    || range.end.col >= MAX_COLS
                {
                    return Err(OpError::InvalidAutoFilter(
                        "Use a valid worksheet range for the auto filter.".into(),
                    ));
                }
                if let Some(filter) = &previous
                    && ((range.start.col != filter.range.start.col
                        && (!filter.filter_columns.is_empty() || filter.has_sort_state))
                        || filter
                            .filter_columns
                            .iter()
                            .any(|col| *col > range.end.col - range.start.col))
                {
                    return Err(OpError::InvalidAutoFilter("Keep the filter's first column and include all filtered columns, or remove the filter before creating a new range.".into()));
                }
                if wb
                    .tables
                    .iter()
                    .any(|table| table.sheet == *sheet && range.overlaps(&table.range))
                {
                    return Err(OpError::InvalidAutoFilter("The auto filter range overlaps a table. Choose a range outside tables; table filters belong to the table part.".into()));
                }
            }
            let name_op = match range {
                Some(range) => {
                    let absolute = |at: CellRef| {
                        let mut at = at;
                        at.abs_row = true;
                        at.abs_col = true;
                        at.to_a1()
                    };
                    Op::SetDefinedName {
                        name: "_xlnm._FilterDatabase".into(),
                        refers_to: format!(
                            "'{}'!{}:{}",
                            current.name.replace('\'', "''"),
                            absolute(range.start),
                            absolute(range.end)
                        ),
                        scope: Some(*sheet),
                        hidden: Some(true),
                    }
                }
                None => Op::DeleteDefinedName {
                    name: "_xlnm._FilterDatabase".into(),
                    scope: Some(*sheet),
                },
            };
            let names_inverse = apply_in_place(wb, &name_op)?;
            let filter = range.map(|range| {
                let mut filter = previous.clone().unwrap_or(xlsx_model::AutoFilter {
                    range,
                    raw_xml: None,
                    source_count: 1,
                    filter_columns: Vec::new(),
                    has_sort_state: false,
                });
                filter.range = range;
                filter
            });
            sheet_mut(wb, *sheet)?.auto_filter = filter;
            let mut inverse = vec![Op::RestoreAutoFilter {
                sheet: *sheet,
                filter: previous,
            }];
            inverse.extend(names_inverse.0);
            Ok(InvertedOp(inverse))
        }
        Op::RestoreAutoFilter { sheet, filter } => {
            let previous =
                std::mem::replace(&mut sheet_mut(wb, *sheet)?.auto_filter, filter.clone());
            Ok(InvertedOp(vec![Op::RestoreAutoFilter {
                sheet: *sheet,
                filter: previous,
            }]))
        }
        Op::SetConditionalFormats { sheet, formats } => {
            let s = sheet_mut(wb, *sheet)?;
            let old = std::mem::replace(&mut s.conditional_formats, formats.clone());
            Ok(InvertedOp(vec![Op::SetConditionalFormats {
                sheet: *sheet,
                formats: old,
            }]))
        }
        Op::RestoreColStyles { sheet, styles } => {
            let s = sheet_mut(wb, *sheet)?;
            let old = std::mem::replace(&mut s.col_styles, styles.clone());
            Ok(InvertedOp(vec![Op::RestoreColStyles {
                sheet: *sheet,
                styles: old,
            }]))
        }
        Op::RestoreRowStyles { sheet, styles } => {
            let s = sheet_mut(wb, *sheet)?;
            let old = std::mem::replace(&mut s.row_styles, styles.clone());
            Ok(InvertedOp(vec![Op::RestoreRowStyles {
                sheet: *sheet,
                styles: old,
            }]))
        }
        Op::SetCharts { sheet, charts } => {
            let sheet_ref = sheet_mut(wb, *sheet)?;
            let old = std::mem::replace(&mut sheet_ref.charts, charts.clone());
            Ok(InvertedOp(vec![Op::SetCharts {
                sheet: *sheet,
                charts: old,
            }]))
        }
        Op::SetChartAnchor {
            sheet,
            frame,
            part,
            from,
            to,
        } => {
            let sheet_ref = sheet_mut(wb, *sheet)?;
            let chart = sheet_ref
                .charts
                .iter_mut()
                .find(|chart| chart.frame_id() == *frame)
                .ok_or_else(|| OpError::ChartNotFound {
                    frame: frame.clone(),
                })?;
            if !chart.is_recorded_frame(part, *from) {
                return Err(OpError::ChartFrameShifted {
                    frame: frame.clone(),
                });
            }
            chart.anchor = *to;
            Ok(InvertedOp(vec![Op::SetChartAnchor {
                sheet: *sheet,
                frame: frame.clone(),
                part: part.clone(),
                from: *to,
                to: *from,
            }]))
        }
        Op::MergeCells { sheet, range } => {
            let range = &CellRange {
                start: CellRef::new(range.start.row, range.start.col),
                end: CellRef::new(range.end.row, range.end.col),
            };
            let s = sheet_mut(wb, *sheet)?;
            let replaced = s
                .merges
                .iter()
                .copied()
                .filter(|merged| ranges_intersect(*merged, *range))
                .collect::<Vec<_>>();
            if replaced.as_slice() == [*range] {
                return Ok(InvertedOp(vec![]));
            }
            s.merges.retain(|merged| !ranges_intersect(*merged, *range));
            s.merges.push(*range);
            let mut inverse = Vec::with_capacity(replaced.len() + 1);
            inverse.push(Op::UnmergeCells {
                sheet: *sheet,
                range: *range,
            });
            inverse.extend(replaced.into_iter().map(|range| Op::MergeCells {
                sheet: *sheet,
                range,
            }));
            Ok(InvertedOp(inverse))
        }
        Op::UnmergeCells { sheet, range } => {
            let range = &CellRange {
                start: CellRef::new(range.start.row, range.start.col),
                end: CellRef::new(range.end.row, range.end.col),
            };
            let s = sheet_mut(wb, *sheet)?;
            match s.merges.iter().position(|m| m == range) {
                Some(pos) => {
                    s.merges.remove(pos);
                    Ok(InvertedOp(vec![Op::MergeCells {
                        sheet: *sheet,
                        range: *range,
                    }]))
                }
                None => Ok(InvertedOp(vec![])),
            }
        }
        Op::PatchRangeStyle {
            sheet,
            range,
            patch,
        } => apply_range_formats(wb, *sheet, *range, true, None, |format, row, col| {
            patch_cell_format(format, patch, *range, row, col)
        }),
        Op::SetRangeNumberFormat {
            sheet,
            range,
            format,
        } => apply_range_formats(wb, *sheet, *range, true, None, |cell_format, _, _| {
            mutate_number_format(cell_format, format);
            Ok(())
        }),
        Op::ApplyRangeFormat {
            sheet,
            range,
            format,
        } => {
            wb.styles.preserve_style_identity();
            if format.rows == 0
                || format.columns == 0
                || format.formats.len() != (format.rows as usize) * (format.columns as usize)
                || !(format.kept.is_empty() || format.kept.len() == format.formats.len())
            {
                return Err(OpError::InvalidStyle(
                    "captured format dimensions do not match its cells".into(),
                ));
            }
            apply_range_formats(
                wb,
                *sheet,
                *range,
                false,
                Some(format),
                |cell_format, row, col| {
                    let source_row = (row - range.start.row) % format.rows;
                    let source_col = (col - range.start.col) % format.columns;
                    let index = (source_row * format.columns + source_col) as usize;
                    let mut applied = format.formats[index].clone();
                    if let Some(kept) = format.kept.get(index) {
                        kept.keep(cell_format, &mut applied);
                    }
                    *cell_format = applied;
                    Ok(())
                },
            )
        }
        Op::AddSheet { index, name } => {
            let idx = (*index).min(wb.sheets.len());
            wb.sheets.insert(idx, Sheet::new(name.clone()));
            shift_defined_name_scopes(wb, idx);
            Ok(InvertedOp(vec![Op::RemoveSheet { index: idx }]))
        }
        Op::RemoveSheet { index } => remove_sheet(wb, *index),
        Op::MoveSheet { from, to } => {
            if *from >= wb.sheets.len() || *to >= wb.sheets.len() {
                return Err(OpError::SheetIndexOutOfRange((*from).max(*to)));
            }
            let sheet = wb.sheets.remove(*from);
            wb.sheets.insert(*to, sheet);
            let remap = |old: SheetId| {
                let index = old.0 as usize;
                SheetId(if index == *from {
                    *to
                } else if from < to && index > *from && index <= *to {
                    index - 1
                } else if from > to && index >= *to && index < *from {
                    index + 1
                } else {
                    index
                } as u32)
            };
            for name in &mut wb.defined_names {
                name.local_sheet = name.local_sheet.map(remap);
            }
            for table in wb.tables.iter_mut() {
                table.sheet = remap(table.sheet);
            }
            Ok(InvertedOp(vec![Op::MoveSheet {
                from: *to,
                to: *from,
            }]))
        }
        Op::RenameSheet { sheet, name } => {
            let old = wb
                .sheet(*sheet)
                .ok_or(OpError::SheetNotFound(*sheet))?
                .name
                .clone();
            let formulas = rename_sheet_references(wb, &old, name)?;
            let hyperlinks = rename_hyperlink_locations(wb, &old, name);
            let defined_names = rename_defined_names(wb, &old, name);
            let charts = rename_chart_refs(wb, &old, name);
            sheet_mut(wb, *sheet)?.name = name.clone();
            let mut inverse = vec![Op::RestoreSheet {
                sheet: *sheet,
                name: old,
                formulas,
            }];
            inverse.extend(hyperlinks);
            inverse.extend(defined_names);
            inverse.extend(charts);
            Ok(InvertedOp(inverse))
        }
        Op::RestoreSheet {
            sheet,
            name,
            formulas,
        } => {
            let old_name = wb
                .sheet(*sheet)
                .ok_or(OpError::SheetNotFound(*sheet))?
                .name
                .clone();
            let mut old_formulas = Vec::with_capacity(formulas.len());
            for (formula_sheet, cell, _) in formulas {
                let sheet_ref = wb
                    .sheet(*formula_sheet)
                    .ok_or(OpError::SheetNotFound(*formula_sheet))?;
                let state = sheet_ref
                    .cell(*cell)
                    .map(CellState::from)
                    .unwrap_or_default();
                old_formulas.push((*formula_sheet, *cell, state));
            }
            let defined_names = rename_defined_names(wb, &old_name, name);
            let charts = rename_chart_refs(wb, &old_name, name);
            sheet_mut(wb, *sheet)?.name = name.clone();
            for (formula_sheet, cell, state) in formulas {
                sheet_mut(wb, *formula_sheet)?.set_cell(*cell, state.clone().into());
            }
            let mut inverse = vec![Op::RestoreSheet {
                sheet: *sheet,
                name: old_name,
                formulas: old_formulas,
            }];
            inverse.extend(defined_names);
            inverse.extend(charts);
            Ok(InvertedOp(inverse))
        }
        Op::SetDefinedName {
            name,
            refers_to,
            scope,
            hidden,
        } => {
            let existing = defined_name_index(wb, name, *scope)?;
            if existing.is_none() && wb.defined_names.len() >= MAX_DEFINED_NAMES {
                return Err(OpError::TooManyDefinedNames);
            }
            if existing.is_none_or(|index| {
                let defined = &wb.defined_names[index];
                defined.formula != *refers_to
                    || hidden.is_some_and(|hidden| hidden != defined.hidden)
            }) {
                validate_defined_name(wb, name, *scope, existing.is_none())?;
            }
            if existing.is_none_or(|index| wb.defined_names[index].formula != *refers_to) {
                validate_defined_name_formula(refers_to)?;
            }
            let previous = wb.defined_names.clone();
            if let Some(index) = existing {
                let defined = &mut wb.defined_names[index];
                defined.formula = refers_to.clone();
                if let Some(hidden) = hidden {
                    defined.hidden = *hidden;
                }
            } else {
                wb.defined_names.push(DefinedName {
                    name: name.clone(),
                    formula: refers_to.clone(),
                    local_sheet: *scope,
                    hidden: hidden.unwrap_or(false),
                });
            }
            Ok(InvertedOp(vec![Op::SetDefinedNames {
                defined_names: previous,
            }]))
        }
        Op::DeleteDefinedName { name, scope } => {
            let existing = defined_name_index(wb, name, *scope)?;
            let previous = wb.defined_names.clone();
            if let Some(index) = existing {
                wb.defined_names.remove(index);
            }
            Ok(InvertedOp(vec![Op::SetDefinedNames {
                defined_names: previous,
            }]))
        }
        Op::SetDefinedNames { defined_names } => {
            if defined_names.len() > MAX_DEFINED_NAMES {
                return Err(OpError::TooManyDefinedNames);
            }
            let previous = std::mem::replace(&mut wb.defined_names, defined_names.clone());
            Ok(InvertedOp(vec![Op::SetDefinedNames {
                defined_names: previous,
            }]))
        }
        Op::InsertRows { sheet, at, count } => insert_rows(wb, *sheet, *at, *count, op),
        Op::DeleteRows { sheet, at, count } => delete_rows(wb, *sheet, *at, *count, op),
        Op::InsertCols { sheet, at, count } => insert_cols(wb, *sheet, *at, *count, op),
        Op::DeleteCols { sheet, at, count } => delete_cols(wb, *sheet, *at, *count, op),
    }
}

fn defined_name_index(
    wb: &Workbook,
    name: &str,
    scope: Option<SheetId>,
) -> Result<Option<usize>, OpError> {
    if let Some(sheet) = scope {
        wb.sheet(sheet).ok_or(OpError::SheetNotFound(sheet))?;
    }
    let mut matches = wb.defined_names.iter().enumerate().filter(|(_, defined)| {
        defined.local_sheet == scope && defined.name.eq_ignore_ascii_case(name)
    });
    let index = matches.next().map(|(index, _)| index);
    if matches.next().is_some() {
        return Err(OpError::InvalidDefinedName(
            "ambiguous defined name in this scope".into(),
        ));
    }
    Ok(index)
}

fn validate_defined_name(
    wb: &Workbook,
    name: &str,
    scope: Option<SheetId>,
    creating: bool,
) -> Result<(), OpError> {
    let invalid = || OpError::InvalidDefinedName("invalid or reserved defined name".into());
    let mut chars = name.chars();
    if name.encode_utf16().count() > 255
        || !chars
            .next()
            .is_some_and(|c| c.is_alphabetic() || c == '_' || c == '\\')
        || !chars.all(|c| c.is_alphabetic() || c.is_ascii_digit() || c == '_' || c == '.')
    {
        return Err(invalid());
    }
    let upper = name.to_ascii_uppercase();
    let r1c1 = upper
        .strip_prefix('R')
        .and_then(|rest| rest.split_once('C'))
        .is_some_and(|(row, col)| {
            row.bytes().all(|b| b.is_ascii_digit()) && col.bytes().all(|b| b.is_ascii_digit())
        });
    if upper == "R" || upper == "C" || CellRef::parse_a1(&upper).is_ok() || r1c1 {
        return Err(invalid());
    }
    if creating
        && upper.starts_with("_XLNM.")
        && !(scope.is_some()
            && matches!(
                upper.as_str(),
                "_XLNM.PRINT_AREA" | "_XLNM.PRINT_TITLES" | "_XLNM._FILTERDATABASE"
            ))
    {
        return Err(invalid());
    }
    if wb
        .tables
        .iter()
        .any(|table| table.name.eq_ignore_ascii_case(name))
    {
        return Err(OpError::InvalidDefinedName(
            "defined name conflicts with a table name".into(),
        ));
    }
    Ok(())
}

fn validate_defined_name_formula(source: &str) -> Result<(), OpError> {
    if source.is_empty() || source.starts_with('=')
        || source.len() > xlsx_calc::lexer::MAX_FORMULA_BYTES
        || !source.chars().all(|c| matches!(c, '\u{9}' | '\u{a}' | '\u{d}' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}'))
    {
        return Err(OpError::InvalidDefinedName("invalid defined name formula".into()));
    }
    xlsx_calc::parse_formula(source)
        .map(|_| ())
        .map_err(|error| {
            OpError::InvalidDefinedName(format!("invalid defined name formula: {error}"))
        })
}

fn apply_atomically(
    wb: &mut Workbook,
    apply: impl FnOnce(&mut Workbook) -> Result<InvertedOp, OpError>,
) -> Result<InvertedOp, OpError> {
    let mut next = wb.clone();
    let inverse = apply(&mut next)?;
    *wb = next;
    Ok(inverse)
}

fn apply_range_formats(
    wb: &mut Workbook,
    sheet: SheetId,
    range: CellRange,
    partial: bool,
    captured: Option<&crate::formatting::CapturedFormat>,
    mut update: impl FnMut(&mut xlsx_model::CellFormat, u32, u32) -> Result<(), OpError>,
) -> Result<InvertedOp, OpError> {
    if wb.sheet(sheet).is_none() {
        return Err(OpError::SheetNotFound(sheet));
    }
    let marks = wb.styles.pool_marks();
    let rows = range
        .end
        .row
        .checked_sub(range.start.row)
        .map_or(0, |rows| u64::from(rows) + 1);
    let cols = range
        .end
        .col
        .checked_sub(range.start.col)
        .map_or(0, |cols| u64::from(cols) + 1);
    let range_cells = rows.saturating_mul(cols);
    let mut staged = Vec::with_capacity(usize::try_from(range_cells).unwrap_or(0));
    let mut resolved_formats = BTreeMap::new();
    {
        let sheet_ref = &wb.sheets[sheet.0 as usize];
        let styles = &mut wb.styles;
        let plain_default = styles.cell_format(Some(0)) == xlsx_model::CellFormat::default();
        let mut plain = None;
        let mut defaults = xlsx_model::workbook::ColumnDefaults::new(sheet_ref);
        for row in range.start.row..=range.end.row {
            for col in range.start.col..=range.end.col {
                let at = CellRef::new(row, col);
                let stored = sheet_ref.cell(at).map(|cell| cell.style);
                let old_style = match stored {
                    Some(style) => style,
                    None => defaults.inherited_style(at),
                };
                let before = styles.cell_format(Some(old_style.unwrap_or(0)));
                let mut format = before.clone();
                if let Err(error) = update(&mut format, at.row, at.col) {
                    styles.restore_pools(marks);
                    return Err(error);
                }
                if let Some(captured) = captured {
                    let source_row = (row - range.start.row) % captured.rows;
                    let source_col = (col - range.start.col) % captured.columns;
                    let index = (source_row * captured.columns + source_col) as usize;
                    if let Some(source) = captured.source_formats.get(index) {
                        let matches = |style| {
                            crate::formatting::PreservedCellFormat::from_style(styles, style)
                                == *source
                        };
                        let selected = if source.format == format {
                            resolved_formats.get(&index).copied().or_else(|| {
                                captured
                                    .source_styles
                                    .get(index)
                                    .copied()
                                    .filter(|style| {
                                        style.is_none_or(|index| styles.xf(index).is_some())
                                            && matches(*style)
                                    })
                                    .or_else(|| {
                                        (0..styles.cell_xfs.len() as u32)
                                            .find(|&index| matches(Some(index)))
                                            .map(Some)
                                    })
                            })
                        } else {
                            None
                        };
                        let selected = match selected {
                            Some(style) => Ok(style),
                            None if format == xlsx_model::CellFormat::default() => Ok(None),
                            None => match source.derived_from {
                                Some(base) if styles.xf(base).is_some() => {
                                    styles.intern_derived_cell_format(base, &format).map(Some)
                                }
                                _ => styles.intern_cell_format(&format),
                            },
                        };
                        match selected {
                            Ok(style) => {
                                if source.format == format {
                                    resolved_formats.insert(index, style);
                                }
                                let style = style.or_else(|| {
                                    if !plain_default {
                                        Some(*plain.get_or_insert_with(|| styles.plain_edit_xf()))
                                    } else {
                                        defaults
                                            .inherited_style(at)
                                            .is_some_and(|inherited| inherited != 0)
                                            .then_some(0)
                                    }
                                });
                                staged.push(style);
                            }
                            Err(_) => {
                                styles.restore_pools(marks);
                                return Err(OpError::NumFmtTableFull);
                            }
                        }
                        continue;
                    }
                }
                if format == before {
                    staged.push(stored.flatten());
                    continue;
                }
                let style = if format == xlsx_model::CellFormat::default() {
                    Ok(None)
                } else if partial {
                    styles
                        .intern_derived_cell_format(old_style.unwrap_or(0), &format)
                        .map(Some)
                } else {
                    styles.intern_cell_format(&format)
                };
                match style {
                    Ok(Some(style)) => staged.push(Some(style)),
                    Ok(None) if !plain_default => {
                        staged.push(Some(*plain.get_or_insert_with(|| styles.plain_edit_xf())));
                    }
                    Ok(None) => staged.push(
                        defaults
                            .inherited_style(at)
                            .is_some_and(|inherited| inherited != 0)
                            .then_some(0),
                    ),
                    Err(_) => {
                        styles.restore_pools(marks);
                        return Err(OpError::NumFmtTableFull);
                    }
                }
            }
        }
    }
    let mut inverse = Vec::with_capacity(staged.len());
    let sheet_ref = wb.sheet_mut(sheet).ok_or(OpError::SheetNotFound(sheet))?;
    let mut staged = staged.into_iter();
    for row in range.start.row..=range.end.row {
        for col in range.start.col..=range.end.col {
            let at = CellRef::new(row, col);
            let style = staged.next().expect("stage covers the range");
            let (old, remove) = match sheet_ref.cell_mut(at) {
                Some(cell) if style != cell.style => {
                    let old = CellState::from(&*cell);
                    cell.style = style;
                    (old, *cell == Cell::default())
                }
                Some(_) => continue,
                None if style.is_some() => {
                    sheet_ref.set_cell(
                        at,
                        Cell {
                            style,
                            ..Cell::default()
                        },
                    );
                    (CellState::default(), false)
                }
                None => continue,
            };
            if remove {
                sheet_ref.set_cell(at, Cell::default());
            }
            inverse.push(Op::SetCell {
                sheet,
                at,
                cell: old,
            });
        }
    }
    Ok(InvertedOp(inverse))
}

/// apply a sequence of ops, returning the combined inverse (per-op inverses
/// concatenated in reverse order).
pub fn apply_ops(wb: &mut Workbook, ops: &[Op]) -> Result<Vec<Op>, OpError> {
    if let [op @ Op::SetCell { .. }] = ops {
        #[cfg(test)]
        if FORCE_CLONE_ROLLBACK.get() {
            return apply_ops_cloned(wb, ops);
        }
        return apply_in_place(wb, op).map(|inverse| inverse.0);
    }
    apply_ops_cloned(wb, ops)
}

#[cfg(test)]
thread_local! {
    static FORCE_CLONE_ROLLBACK: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
}

fn apply_ops_cloned(wb: &mut Workbook, ops: &[Op]) -> Result<Vec<Op>, OpError> {
    let mut next = wb.clone();
    let inverse = apply_ops_in_place(&mut next, ops)?;
    *wb = next;
    Ok(inverse)
}

fn apply_ops_in_place(wb: &mut Workbook, ops: &[Op]) -> Result<Vec<Op>, OpError> {
    let mut per_op: Vec<Vec<Op>> = Vec::with_capacity(ops.len());
    for op in ops {
        per_op.push(apply_in_place(wb, op)?.0);
    }
    let mut inverse = Vec::new();
    for chunk in per_op.into_iter().rev() {
        inverse.extend(chunk);
    }
    Ok(inverse)
}

fn ranges_intersect(left: CellRange, right: CellRange) -> bool {
    left.start.row <= right.end.row
        && left.end.row >= right.start.row
        && left.start.col <= right.end.col
        && left.end.col >= right.start.col
}

/// remap an address through a structural op; `None` when it falls inside a
/// deleted span. non-structural ops leave the address unchanged.
pub fn remap_ref(at: CellRef, structural_op: &Op) -> Option<CellRef> {
    match *structural_op {
        Op::InsertRows { at: p, count, .. } => {
            shift_row(at, |r| insert_axis(r, p, count, MAX_ROWS))
        }
        Op::DeleteRows { at: p, count, .. } => shift_row(at, |r| delete_axis(r, p, count)),
        Op::InsertCols { at: p, count, .. } => {
            shift_col(at, |c| insert_axis(c, p, count, MAX_COLS))
        }
        Op::DeleteCols { at: p, count, .. } => shift_col(at, |c| delete_axis(c, p, count)),
        _ => Some(at),
    }
}

/// remap both corners of a range; `None` if either corner is deleted.
fn remap_range(range: CellRange, op: &Op) -> Option<CellRange> {
    let start = remap_ref(range.start, op)?;
    let end = remap_ref(range.end, op)?;
    Some(CellRange::new(start, end))
}

/// index shift under an insert at `p` of `count`. indices at or after `p` move
/// up by `count`; anything pushed past `limit` is dropped (`None`).
fn insert_axis(idx: u32, p: u32, count: u32, limit: u32) -> Option<u32> {
    if idx < p {
        return Some(idx);
    }
    let shifted = u64::from(idx) + u64::from(count);
    if shifted > u64::from(limit - 1) {
        None
    } else {
        Some(shifted as u32)
    }
}

/// index shift under a delete at `p` of `count`. indices inside `[p, p+count)`
/// are deleted (`None`); later indices move down by `count`.
fn delete_axis(idx: u32, p: u32, count: u32) -> Option<u32> {
    if idx < p {
        Some(idx)
    } else if idx < p.saturating_add(count) {
        None
    } else {
        Some(idx - count)
    }
}

fn shift_row(mut at: CellRef, f: impl Fn(u32) -> Option<u32>) -> Option<CellRef> {
    at.row = f(at.row)?;
    Some(at)
}

fn shift_col(mut at: CellRef, f: impl Fn(u32) -> Option<u32>) -> Option<CellRef> {
    at.col = f(at.col)?;
    Some(at)
}

fn insert_rows(
    wb: &mut Workbook,
    sheet: SheetId,
    at: RowId,
    count: u32,
    op: &Op,
) -> Result<InvertedOp, OpError> {
    wb.sheet(sheet).ok_or(OpError::SheetNotFound(sheet))?;
    let restores = remap_formulas(wb, op)?;
    let defined_name_restore = remap_defined_names(wb, op)?;
    let chart_restores = remap_charts(wb, op)?;
    let hyperlink_restores = remap_hyperlink_locations(wb, op);
    let s = sheet_mut(wb, sheet)?;
    let old_hyperlinks = s.hyperlinks.clone();
    let dropped = shift_cells(s, op);
    let old_row_styles = s.row_styles.clone();
    let inherited = at
        .checked_sub(1)
        .and_then(|row| s.row_styles.get(&row))
        .copied();
    let retained = s
        .row_styles
        .keys()
        .filter(|&&row| insert_axis(row, at, count, MAX_ROWS).is_some())
        .count();
    let added = if inherited.is_some() {
        count.min(MAX_ROWS.saturating_sub(at)) as usize
    } else {
        0
    };
    if retained.saturating_add(added) > xlsx_model::workbook::MAX_ROW_STYLES {
        return Err(OpError::InvalidStyle("row style count exceeded cap".into()));
    }
    s.row_styles = std::mem::take(&mut s.row_styles)
        .into_iter()
        .filter_map(|(row, style)| insert_axis(row, at, count, MAX_ROWS).map(|row| (row, style)))
        .collect();
    if let Some(style) = inherited {
        for row in at..at.saturating_add(count).min(MAX_ROWS) {
            s.row_styles.insert(row, style);
        }
    }
    shift_row_heights_up(s, at, count);
    remap_merges_keep(s, op);
    remap_hyperlinks(s, op);
    let auto_filter_restore = remap_auto_filter(s, sheet, op);

    let mut inv = vec![Op::DeleteRows { sheet, at, count }];
    if !old_row_styles.is_empty() {
        inv.push(Op::RestoreRowStyles {
            sheet,
            styles: old_row_styles,
        });
    }
    for (r, c) in dropped {
        inv.push(Op::SetCell {
            sheet,
            at: r,
            cell: CellState::from(&c),
        });
    }
    inv.push(Op::SetHyperlinks {
        sheet,
        hyperlinks: old_hyperlinks,
    });
    inv.extend(auto_filter_restore);
    inv.extend(restores);
    inv.extend(defined_name_restore);
    inv.extend(chart_restores);
    inv.extend(hyperlink_restores);
    Ok(InvertedOp(inv))
}

fn delete_rows(
    wb: &mut Workbook,
    sheet: SheetId,
    at: RowId,
    count: u32,
    op: &Op,
) -> Result<InvertedOp, OpError> {
    wb.sheet(sheet).ok_or(OpError::SheetNotFound(sheet))?;
    let restores = remap_formulas(wb, op)?;
    let defined_name_restore = remap_defined_names(wb, op)?;
    let chart_restores = remap_charts(wb, op)?;
    let hyperlink_restores = remap_hyperlink_locations(wb, op);
    let s = sheet_mut(wb, sheet)?;
    let old_hyperlinks = s.hyperlinks.clone();
    let deleted = shift_cells(s, op);
    let old_row_styles = s.row_styles.clone();
    s.row_styles = std::mem::take(&mut s.row_styles)
        .into_iter()
        .filter_map(|(row, style)| delete_axis(row, at, count).map(|row| (row, style)))
        .collect();
    let dropped_heights = shift_row_heights_down(s, at, count);
    let dropped_hidden = shift_number_map_down(&mut s.hidden_row_heights, at, count)
        .into_iter()
        .collect::<BTreeMap<_, _>>();
    let dropped_merges = remap_merges_drop(s, op);
    remap_hyperlinks(s, op);
    let auto_filter_restore = remap_auto_filter(s, sheet, op);

    let mut inv = Vec::new();
    if !old_row_styles.is_empty() {
        // the restore after the insert sets the map, so the insert must neither
        // inherit nor count rows against the cap
        inv.push(Op::RestoreRowStyles {
            sheet,
            styles: BTreeMap::new(),
        });
    }
    inv.push(Op::InsertRows { sheet, at, count });
    if !old_row_styles.is_empty() {
        inv.push(Op::RestoreRowStyles {
            sheet,
            styles: old_row_styles,
        });
    }
    for (r, c) in deleted {
        inv.push(Op::SetCell {
            sheet,
            at: r,
            cell: CellState::from(&c),
        });
    }
    for (row, h) in dropped_heights {
        inv.push(Op::RestoreRowDimension {
            sheet,
            row,
            height: Some(h),
            hidden_height: dropped_hidden.get(&row).copied(),
        });
    }
    for range in dropped_merges {
        inv.push(Op::MergeCells { sheet, range });
    }
    inv.push(Op::SetHyperlinks {
        sheet,
        hyperlinks: old_hyperlinks,
    });
    inv.extend(auto_filter_restore);
    inv.extend(restores);
    inv.extend(defined_name_restore);
    inv.extend(chart_restores);
    inv.extend(hyperlink_restores);
    Ok(InvertedOp(inv))
}

fn insert_cols(
    wb: &mut Workbook,
    sheet: SheetId,
    at: ColId,
    count: u32,
    op: &Op,
) -> Result<InvertedOp, OpError> {
    wb.sheet(sheet).ok_or(OpError::SheetNotFound(sheet))?;
    let restores = remap_formulas(wb, op)?;
    let defined_name_restore = remap_defined_names(wb, op)?;
    let chart_restores = remap_charts(wb, op)?;
    let hyperlink_restores = remap_hyperlink_locations(wb, op);
    let s = sheet_mut(wb, sheet)?;
    let old_hyperlinks = s.hyperlinks.clone();
    let old_col_styles = s.col_styles.clone();
    let dropped = shift_cells(s, op);
    shift_col_widths_up(s, at, count);
    shift_col_styles_up(s, at, count);
    remap_merges_keep(s, op);
    remap_hyperlinks(s, op);
    let auto_filter_restore = remap_auto_filter(s, sheet, op);

    let mut inv = vec![Op::DeleteCols { sheet, at, count }];
    if !old_col_styles.is_empty() {
        inv.push(Op::RestoreColStyles {
            sheet,
            styles: old_col_styles,
        });
    }
    for (r, c) in dropped {
        inv.push(Op::SetCell {
            sheet,
            at: r,
            cell: CellState::from(&c),
        });
    }
    inv.push(Op::SetHyperlinks {
        sheet,
        hyperlinks: old_hyperlinks,
    });
    inv.extend(auto_filter_restore);
    inv.extend(restores);
    inv.extend(defined_name_restore);
    inv.extend(chart_restores);
    inv.extend(hyperlink_restores);
    Ok(InvertedOp(inv))
}

fn delete_cols(
    wb: &mut Workbook,
    sheet: SheetId,
    at: ColId,
    count: u32,
    op: &Op,
) -> Result<InvertedOp, OpError> {
    wb.sheet(sheet).ok_or(OpError::SheetNotFound(sheet))?;
    let restores = remap_formulas(wb, op)?;
    let defined_name_restore = remap_defined_names(wb, op)?;
    let chart_restores = remap_charts(wb, op)?;
    let hyperlink_restores = remap_hyperlink_locations(wb, op);
    let s = sheet_mut(wb, sheet)?;
    let old_hyperlinks = s.hyperlinks.clone();
    let old_col_styles = s.col_styles.clone();
    let deleted = shift_cells(s, op);
    let dropped_widths = shift_col_widths_down(s, at, count);
    let dropped_hidden = shift_number_map_down(&mut s.hidden_col_widths, at, count)
        .into_iter()
        .collect::<BTreeMap<_, _>>();
    shift_col_styles_down(s, at, count);
    let dropped_merges = remap_merges_drop(s, op);
    remap_hyperlinks(s, op);
    let auto_filter_restore = remap_auto_filter(s, sheet, op);

    let mut inv = vec![Op::InsertCols { sheet, at, count }];
    if !old_col_styles.is_empty() {
        inv.push(Op::RestoreColStyles {
            sheet,
            styles: old_col_styles,
        });
    }
    for (r, c) in deleted {
        inv.push(Op::SetCell {
            sheet,
            at: r,
            cell: CellState::from(&c),
        });
    }
    for (col, w) in dropped_widths {
        inv.push(Op::RestoreColDimension {
            sheet,
            col,
            width: Some(w),
            hidden_width: dropped_hidden.get(&col).copied(),
        });
    }
    for range in dropped_merges {
        inv.push(Op::MergeCells { sheet, range });
    }
    inv.push(Op::SetHyperlinks {
        sheet,
        hyperlinks: old_hyperlinks,
    });
    inv.extend(auto_filter_restore);
    inv.extend(restores);
    inv.extend(defined_name_restore);
    inv.extend(chart_restores);
    inv.extend(hyperlink_restores);
    Ok(InvertedOp(inv))
}

/// remap every occupied cell through `op` in place. returns the cells whose
/// address was dropped, for the inverse.
fn shift_cells(s: &mut Sheet, op: &Op) -> Vec<(CellRef, Cell)> {
    s.remap_cells(|at| remap_ref(at, op))
}

fn shift_row_heights_up(s: &mut Sheet, at: RowId, count: u32) {
    shift_number_map_up(&mut s.row_heights, at, count);
    shift_number_map_up(&mut s.hidden_row_heights, at, count);
}

fn shift_number_map_up(map: &mut BTreeMap<u32, f64>, at: u32, count: u32) {
    let shifted: BTreeMap<u32, f64> = map
        .iter()
        .map(|(&row, &h)| {
            (
                if row >= at {
                    row.saturating_add(count)
                } else {
                    row
                },
                h,
            )
        })
        .collect();
    *map = shifted;
}

fn shift_row_heights_down(s: &mut Sheet, at: RowId, count: u32) -> Vec<(RowId, f64)> {
    shift_number_map_down(&mut s.row_heights, at, count)
}

fn shift_number_map_down(map: &mut BTreeMap<u32, f64>, at: u32, count: u32) -> Vec<(u32, f64)> {
    let mut kept = BTreeMap::new();
    let mut dropped = Vec::new();
    for (&row, &h) in map.iter() {
        if row < at {
            kept.insert(row, h);
        } else if row >= at.saturating_add(count) {
            kept.insert(row - count, h);
        } else {
            dropped.push((row, h));
        }
    }
    *map = kept;
    dropped
}

fn shift_col_widths_up(s: &mut Sheet, at: ColId, count: u32) {
    shift_number_map_up(&mut s.col_widths, at, count);
    shift_number_map_up(&mut s.hidden_col_widths, at, count);
}

fn shift_col_widths_down(s: &mut Sheet, at: ColId, count: u32) -> Vec<(ColId, f64)> {
    shift_number_map_down(&mut s.col_widths, at, count)
}

fn shift_col_styles_up(s: &mut Sheet, at: ColId, count: u32) {
    if count == 0 {
        return;
    }
    let inherited = at.checked_sub(1).and_then(|col| s.col_style(col));
    s.col_styles.retain_mut(|run| {
        if run.last < at {
            return true;
        }
        if run.first >= at {
            let Some(first) = insert_axis(run.first, at, count, MAX_COLS) else {
                return false;
            };
            run.first = first;
        }
        run.last = run.last.saturating_add(count).min(MAX_COLS - 1);
        true
    });
    if let Some(xf) = inherited
        && at < MAX_COLS
        && s.col_style(at) != Some(xf)
    {
        s.col_styles.push(xlsx_model::ColStyle {
            first: at,
            last: at.saturating_add(count - 1).min(MAX_COLS - 1),
            xf,
        });
    }
}

fn shift_col_styles_down(s: &mut Sheet, at: ColId, count: u32) {
    let end = at.saturating_add(count);
    s.col_styles.retain_mut(|run| {
        if run.first >= at && run.last < end {
            return false;
        }
        if run.first >= end {
            run.first -= count;
        } else if run.first >= at {
            run.first = at;
        }
        if run.last >= end {
            run.last -= count;
        } else if run.last >= at {
            run.last = at - 1;
        }
        true
    });
}

/// remap merges under an insert (no corner is ever deleted).
fn remap_merges_keep(s: &mut Sheet, op: &Op) {
    let remapped: Vec<CellRange> = s
        .merges
        .iter()
        .filter_map(|m| remap_range(*m, op))
        .collect();
    s.merges = remapped;
}

/// remap merges under a delete; any merge with a deleted corner is dropped and
/// returned for the inverse to restore.
fn remap_merges_drop(s: &mut Sheet, op: &Op) -> Vec<CellRange> {
    let mut kept = Vec::new();
    let mut dropped = Vec::new();
    for m in &s.merges {
        match remap_range(*m, op) {
            Some(nm) => kept.push(nm),
            None => dropped.push(*m),
        }
    }
    s.merges = kept;
    dropped
}

fn remap_hyperlinks(sheet: &mut Sheet, op: &Op) {
    sheet.hyperlinks = sheet
        .hyperlinks
        .drain(..)
        .filter_map(|mut hyperlink| {
            hyperlink.range = remap_hyperlink_range(hyperlink.range, op)?;
            Some(hyperlink)
        })
        .collect();
}

fn rename_hyperlink_locations(wb: &mut Workbook, old_name: &str, new_name: &str) -> Vec<Op> {
    let mut restores = Vec::new();
    for (index, sheet) in wb.sheets.iter_mut().enumerate() {
        let mut changed = false;
        let mut hyperlinks = sheet.hyperlinks.clone();
        for hyperlink in &mut hyperlinks {
            let Some(location) = &hyperlink.location else {
                continue;
            };
            let rewritten = rename_hyperlink_location(location, old_name, new_name);
            if rewritten != *location {
                hyperlink.location = Some(rewritten);
                changed = true;
            }
        }
        if changed {
            let old = std::mem::replace(&mut sheet.hyperlinks, hyperlinks);
            restores.push(Op::SetHyperlinks {
                sheet: SheetId(index as u32),
                hyperlinks: old,
            });
        }
    }
    restores
}

fn remove_sheet(wb: &mut Workbook, index: usize) -> Result<InvertedOp, OpError> {
    if index >= wb.sheets.len() {
        return Err(OpError::SheetIndexOutOfRange(index));
    }
    let order_before = wb
        .sheets
        .iter()
        .map(|sheet| sheet.name.clone())
        .collect::<Vec<_>>();
    let removed = wb.sheets.remove(index);
    let previous_defined_names = drop_defined_name_scopes(wb, index);
    let chart_restores = strand_chart_refs(wb, &removed.name, &order_before)
        .into_iter()
        .map(|restore| rebase_chart_restore(restore, index))
        .collect::<Vec<_>>();
    let sheet = SheetId(index as u32);
    let mut inv = vec![Op::AddSheet {
        index,
        name: removed.name.clone(),
    }];
    for (at, cell) in removed.iter_cells() {
        inv.push(Op::SetCell {
            sheet,
            at,
            cell: CellState::from(cell),
        });
    }
    for range in &removed.merges {
        inv.push(Op::MergeCells {
            sheet,
            range: *range,
        });
    }
    for (&col, &w) in &removed.col_widths {
        inv.push(Op::RestoreColDimension {
            sheet,
            col,
            width: Some(w),
            hidden_width: removed.hidden_col_widths.get(&col).copied(),
        });
    }
    for (&row, &h) in &removed.row_heights {
        inv.push(Op::RestoreRowDimension {
            sheet,
            row,
            height: Some(h),
            hidden_height: removed.hidden_row_heights.get(&row).copied(),
        });
    }
    if let Some(pane) = removed.freeze_pane {
        inv.push(Op::SetFreezePane {
            sheet,
            pane: Some(pane),
        });
    }
    if removed.auto_filter.is_some() {
        inv.push(Op::RestoreAutoFilter {
            sheet,
            filter: removed.auto_filter,
        });
    }
    if !removed.hyperlinks.is_empty() {
        inv.push(Op::SetHyperlinks {
            sheet,
            hyperlinks: removed.hyperlinks,
        });
    }
    if !removed.col_styles.is_empty() {
        inv.push(Op::RestoreColStyles {
            sheet,
            styles: removed.col_styles,
        });
    }
    if !removed.row_styles.is_empty() {
        inv.push(Op::RestoreRowStyles {
            sheet,
            styles: removed.row_styles,
        });
    }
    if !removed.charts.is_empty() {
        inv.push(Op::SetCharts {
            sheet,
            charts: removed.charts,
        });
    }
    if let Some(defined_names) = previous_defined_names {
        inv.push(Op::SetDefinedNames { defined_names });
    }
    inv.extend(chart_restores);
    Ok(InvertedOp(inv))
}

/// Rebases chart restores after sheet insertion.
fn rebase_chart_restore(restore: Op, removed_index: usize) -> Op {
    match restore {
        Op::SetCharts {
            sheet: SheetId(at),
            charts,
        } if at as usize >= removed_index => Op::SetCharts {
            sheet: SheetId(at + 1),
            charts,
        },
        other => other,
    }
}

/// Widens sheet scopes at or after an inserted sheet.
fn shift_defined_name_scopes(wb: &mut Workbook, index: usize) {
    for defined in &mut wb.defined_names {
        if let Some(sheet) = defined.local_sheet
            && sheet.0 as usize >= index
        {
            defined.local_sheet = Some(SheetId(sheet.0 + 1));
        }
    }
}

/// Drops names scoped to a removed sheet and narrows the scopes above it,
/// returning the prior list when it changed.
fn drop_defined_name_scopes(wb: &mut Workbook, index: usize) -> Option<Vec<DefinedName>> {
    let previous = wb.defined_names.clone();
    wb.defined_names
        .retain(|defined| defined.local_sheet.map(|sheet| sheet.0 as usize) != Some(index));
    for defined in &mut wb.defined_names {
        if let Some(sheet) = defined.local_sheet
            && sheet.0 as usize > index
        {
            defined.local_sheet = Some(SheetId(sheet.0 - 1));
        }
    }
    (previous != wb.defined_names).then_some(previous)
}

const CELL_IS_OPERATORS: &[&str] = &[
    "equal",
    "notEqual",
    "greaterThan",
    "lessThan",
    "greaterThanOrEqual",
    "lessThanOrEqual",
    "between",
    "notBetween",
];

fn validate_conditional_format(
    kind: &str,
    operator: Option<&str>,
    formulas: &[String],
    format: &DifferentialFormat,
) -> Result<Vec<String>, OpError> {
    let invalid = |message: &str| OpError::InvalidConditionalFormat(message.into());
    let expected = match kind {
        "expression" => {
            if operator.is_some() {
                return Err(invalid("expression rules take no operator"));
            }
            1
        }
        "cellIs" => match operator {
            Some(operator) if CELL_IS_OPERATORS.contains(&operator) => {
                if matches!(operator, "between" | "notBetween") {
                    2
                } else {
                    1
                }
            }
            Some(_) => return Err(invalid("unsupported cellIs operator")),
            None => return Err(invalid("cellIs rules require an operator")),
        },
        _ => return Err(invalid("only cellIs and expression rules can be authored")),
    };
    if formulas.len() != expected {
        return Err(invalid("wrong number of formulas for this rule"));
    }
    let mut cleaned = Vec::with_capacity(formulas.len());
    for formula in formulas {
        let valid_xml = formula.chars().all(|ch| {
            matches!(ch, '\u{9}' | '\u{a}' | '\u{d}' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')
        });
        if !valid_xml {
            return Err(invalid("formulas contain XML 1.0 forbidden characters"));
        }
        let formula = formula.strip_prefix('=').unwrap_or(formula).trim();
        if formula.is_empty() || formula.len() > 8192 {
            return Err(invalid("formulas must be 1 to 8192 characters"));
        }
        cleaned.push(formula.to_owned());
    }
    for color in [&format.font_color, &format.fill].into_iter().flatten() {
        let hex = color.strip_prefix('#').unwrap_or("");
        if hex.len() != 6 || !hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(invalid("colors must be in #RRGGBB format"));
        }
    }
    Ok(cleaned)
}

fn next_conditional_format_id(groups: &[ConditionalFormatting]) -> String {
    let mut next = 0u64;
    for rule in groups.iter().flat_map(|group| &group.rules) {
        if let Some(n) = rule
            .id
            .strip_prefix("cf")
            .and_then(|rest| rest.parse::<u64>().ok())
        {
            next = next.max(n);
        }
    }
    format!("cf{}", next + 1)
}

fn sheet_mut(wb: &mut Workbook, sheet: SheetId) -> Result<&mut Sheet, OpError> {
    wb.sheet_mut(sheet).ok_or(OpError::SheetNotFound(sheet))
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::{CellProvider, CellValue, ColStyle, FreezePane, Hyperlink};

    fn r(a1: &str) -> CellRef {
        CellRef::parse_a1(a1).unwrap()
    }

    fn num(v: f64) -> CellState {
        CellState {
            value: CellValue::Number { value: v },
            ..Default::default()
        }
    }

    fn wb_one_sheet() -> Workbook {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Sheet1"));
        wb
    }

    fn add_cf(formula: &str) -> Op {
        Op::AddConditionalFormat {
            sheet: SheetId(0),
            range: CellRange::parse_a1("A1:A5").unwrap(),
            kind: "expression".into(),
            operator: None,
            formulas: vec![formula.into()],
            format: DifferentialFormat::default(),
        }
    }

    #[test]
    fn authored_conditional_formats_refuse_grid_edits_atomically() {
        for op in [
            Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
            Op::InsertCols {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
            Op::DeleteCols {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        ] {
            let mut wb = wb_one_sheet();
            let before = wb.clone();
            assert!(matches!(
                apply_ops(&mut wb, &[add_cf("A1>0"), op.clone()]),
                Err(OpError::InvalidConditionalFormat(_))
            ));
            assert_eq!(wb, before);
            apply(&mut wb, &add_cf("A1>0")).unwrap();
            let before = wb.clone();
            assert!(matches!(
                apply(&mut wb, &op),
                Err(OpError::InvalidConditionalFormat(_))
            ));
            assert_eq!(wb, before);
        }
    }

    #[test]
    fn authored_conditional_formats_allow_grid_edits_on_other_sheets() {
        let mut wb = wb_one_sheet();
        wb.sheets.push(Sheet::new("Other"));
        apply(&mut wb, &add_cf("A1>0")).unwrap();
        let formats = wb.sheets[0].conditional_formats.clone();
        apply(
            &mut wb,
            &Op::InsertRows {
                sheet: SheetId(1),
                at: 0,
                count: 1,
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].conditional_formats, formats);
    }

    #[test]
    fn conditional_format_formulas_refuse_xml_forbidden_characters_atomically() {
        for ch in (0..=0x1f)
            .filter(|ch| !matches!(ch, 9 | 10 | 13))
            .chain([0xfffe, 0xffff])
        {
            let mut wb = wb_one_sheet();
            let before = wb.clone();
            let formula = format!("{}A1>0", char::from_u32(ch).unwrap());
            assert!(matches!(
                apply(&mut wb, &add_cf(&formula)),
                Err(OpError::InvalidConditionalFormat(_))
            ));
            assert_eq!(wb, before);
            assert!(matches!(
                apply_ops(&mut wb, &[add_cf("A1>0"), add_cf(&formula)]),
                Err(OpError::InvalidConditionalFormat(_))
            ));
            assert_eq!(wb, before);
        }
        let mut wb = wb_one_sheet();
        apply(&mut wb, &add_cf("\t=\"é😀\"<>\"\"\n")).unwrap();
    }

    #[test]
    fn singleton_set_cell_matches_clone_rollback() {
        let states = [
            CellState::default(),
            num(123.0),
            CellState {
                value: CellValue::Text {
                    value: "rich text".into(),
                },
                style: Some(7),
                ..Default::default()
            },
            CellState {
                value: CellValue::Number { value: 42.0 },
                formula: Some("A1+1".into()),
                style: Some(0),
            },
            CellState {
                style: Some(u32::MAX),
                ..Default::default()
            },
        ];
        for before in &states {
            for after in &states {
                for at in [r("B2"), CellRef::new(MAX_ROWS, MAX_COLS)] {
                    let mut initial = wb_one_sheet();
                    initial.sheets.push(Sheet::new("Second"));
                    initial.sheets[1].set_cell(at, before.clone().into());
                    initial.sheets[1].set_array_formula(at, CellRange::new(at, at));
                    let op = Op::SetCell {
                        sheet: SheetId(1),
                        at,
                        cell: after.clone(),
                    };
                    let mut fast = initial.clone();
                    let mut oracle = initial.clone();
                    let result = apply_ops(&mut fast, std::slice::from_ref(&op));
                    FORCE_CLONE_ROLLBACK.set(true);
                    let expected = apply_ops(&mut oracle, std::slice::from_ref(&op));
                    FORCE_CLONE_ROLLBACK.set(false);
                    assert_eq!(result, expected);
                    assert_eq!(fast, oracle);
                    apply_ops(&mut fast, &result.unwrap()).unwrap();
                    assert_eq!(fast, initial);
                }
            }
        }
    }

    #[test]
    fn failing_singleton_set_cell_leaves_model_unchanged() {
        for initial in [Workbook::default(), wb_one_sheet()] {
            for sheet in [SheetId(initial.sheets.len() as u32), SheetId(u32::MAX)] {
                let op = Op::SetCell {
                    sheet,
                    at: r("B2"),
                    cell: num(123.0),
                };
                let mut fast = initial.clone();
                let mut oracle = initial.clone();
                let result = apply_ops(&mut fast, std::slice::from_ref(&op));
                let expected = apply_ops_cloned(&mut oracle, std::slice::from_ref(&op));
                assert_eq!(result, Err(OpError::SheetNotFound(sheet)));
                assert_eq!(result, expected);
                assert_eq!(fast, initial);
                assert_eq!(oracle, initial);
            }
        }
    }

    #[test]
    fn remap_ref_insert_rows_edges() {
        let op = Op::InsertRows {
            sheet: SheetId(0),
            at: 5,
            count: 3,
        };
        // a1 row N is 0-based index N-1, so A5 sits before the insert at 5
        assert_eq!(remap_ref(r("A5"), &op), Some(r("A5")));
        assert_eq!(remap_ref(r("A6"), &op).unwrap().row, 5 + 3);
        assert_eq!(remap_ref(r("A100"), &op).unwrap().row, 99 + 3);
    }

    #[test]
    fn zero_height_visibility_tracks_edits_undo_and_redo() {
        let sheet = SheetId(0);
        let mut initial = wb_one_sheet();
        initial.sheets[0].format.zero_height = true;
        initial.sheets[0].set_cell(CellRef::new(1, 0), num(1.0).into());
        initial.sheets[0].set_cell(
            CellRef::new(3, 0),
            Cell {
                style: Some(0),
                ..Cell::default()
            },
        );
        initial.sheets[0].set_cell(CellRef::new(9, 0), num(9.0).into());
        initial.sheets[0].row_styles.insert(5, 0);
        initial.sheets[0].row_heights.extend([(7, 20.0), (9, 0.0)]);
        let visible = |wb: &Workbook| {
            (0..16)
                .filter(|&row| !wb.sheets[0].row_hidden(row))
                .collect::<Vec<_>>()
        };
        for (op, expected) in [
            (
                Op::SetCell {
                    sheet,
                    at: CellRef::new(0, 0),
                    cell: num(2.0),
                },
                vec![0, 1, 3, 5, 7],
            ),
            (
                Op::SetCell {
                    sheet,
                    at: CellRef::new(1, 0),
                    cell: CellState::default(),
                },
                vec![3, 5, 7],
            ),
            (
                Op::SetCell {
                    sheet,
                    at: CellRef::new(3, 0),
                    cell: CellState::default(),
                },
                vec![1, 5, 7],
            ),
            (
                Op::InsertRows {
                    sheet,
                    at: 1,
                    count: 2,
                },
                vec![3, 5, 7, 9],
            ),
            (
                Op::DeleteRows {
                    sheet,
                    at: 1,
                    count: 2,
                },
                vec![1, 3, 5],
            ),
            (
                Op::InsertCols {
                    sheet,
                    at: 0,
                    count: 1,
                },
                vec![1, 3, 5, 7],
            ),
            (
                Op::DeleteCols {
                    sheet,
                    at: 0,
                    count: 1,
                },
                vec![5, 7],
            ),
            (
                Op::SetRowHeight {
                    sheet,
                    row: 1,
                    height: Some(0.0),
                },
                vec![3, 5, 7],
            ),
            (
                Op::RestoreRowStyles {
                    sheet,
                    styles: BTreeMap::new(),
                },
                vec![1, 3, 7],
            ),
        ] {
            let mut workbook = initial.clone();
            let inverse = apply_ops(&mut workbook, std::slice::from_ref(&op)).unwrap();
            let edited = workbook.clone();
            assert_eq!(visible(&workbook), expected, "{op:?}");
            let redo = apply_ops(&mut workbook, &inverse).unwrap();
            assert_eq!(workbook, initial, "undo {op:?}");
            assert_eq!(visible(&workbook), [1, 3, 5, 7]);
            apply_ops(&mut workbook, &redo).unwrap();
            assert_eq!(workbook, edited, "redo {op:?}");
            assert_eq!(visible(&workbook), expected);
        }
    }

    #[test]
    fn remap_ref_delete_rows_edges() {
        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 5,
            count: 3,
        };
        assert_eq!(remap_ref(r("A5"), &op), Some(r("A5")));
        assert_eq!(remap_ref(CellRef::new(5, 0), &op), None);
        assert_eq!(remap_ref(CellRef::new(7, 0), &op), None);
        assert_eq!(remap_ref(CellRef::new(8, 0), &op), Some(CellRef::new(5, 0)));
    }

    #[test]
    fn remap_ref_cols_and_anchors() {
        let ins = Op::InsertCols {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        };
        let mut anchored = r("C1");
        anchored.abs_col = true;
        let out = remap_ref(anchored, &ins).unwrap();
        assert_eq!(out.col, 3);
        assert!(out.abs_col, "anchor preserved through remap");

        let del = Op::DeleteCols {
            sheet: SheetId(0),
            at: 2,
            count: 2,
        };
        assert_eq!(remap_ref(CellRef::new(0, 2), &del), None);
        assert_eq!(
            remap_ref(CellRef::new(0, 4), &del),
            Some(CellRef::new(0, 2))
        );
    }

    #[test]
    fn remap_ref_ignores_non_structural() {
        let op = Op::SetCell {
            sheet: SheetId(0),
            at: r("A1"),
            cell: num(1.0),
        };
        assert_eq!(remap_ref(r("Z9"), &op), Some(r("Z9")));
    }

    #[test]
    fn set_cell_round_trip() {
        let mut wb = wb_one_sheet();
        let op = Op::SetCell {
            sheet: SheetId(0),
            at: r("B2"),
            cell: num(42.0),
        };
        let inv = apply(&mut wb, &op).unwrap();
        assert_eq!(
            wb.value(SheetId(0), r("B2")),
            CellValue::Number { value: 42.0 }
        );
        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.value(SheetId(0), r("B2")), CellValue::Empty);
        assert!(wb.sheet(SheetId(0)).unwrap().used_range().is_none());
    }

    #[test]
    fn delete_rows_restores_contents_and_merges() {
        let mut wb = wb_one_sheet();
        let s = wb.sheet_mut(SheetId(0)).unwrap();
        s.set_cell(
            r("A1"),
            Cell {
                value: CellValue::Number { value: 1.0 },
                ..Default::default()
            },
        );
        s.set_cell(
            r("A6"),
            Cell {
                value: CellValue::Number { value: 6.0 },
                ..Default::default()
            },
        );
        s.merges.push(CellRange::parse_a1("A6:B6").unwrap());
        s.hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("A6:B6").unwrap(),
            external_target: Some("https://example.com".into()),
            location: None,
            tooltip: None,
            display: None,
        });
        s.row_heights.insert(5, 30.0);
        let before = wb.clone();

        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 2,
            count: 2,
        };
        let inv = apply(&mut wb, &op).unwrap();
        assert_eq!(
            wb.value(SheetId(0), r("A4")),
            CellValue::Number { value: 6.0 }
        );

        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        let after = &wb.sheets[0];
        let orig = &before.sheets[0];
        assert_eq!(after.used_range(), orig.used_range());
        assert_eq!(
            wb.value(SheetId(0), r("A6")),
            CellValue::Number { value: 6.0 }
        );
        assert_eq!(after.merges, orig.merges);
        assert_eq!(after.hyperlinks, orig.hyperlinks);
        assert_eq!(after.row_heights, orig.row_heights);
    }

    #[test]
    fn insert_rows_round_trip() {
        let mut wb = wb_one_sheet();
        wb.sheet_mut(SheetId(0)).unwrap().set_cell(
            r("A3"),
            Cell {
                value: CellValue::Text { value: "x".into() },
                ..Default::default()
            },
        );
        wb.sheet_mut(SheetId(0))
            .unwrap()
            .hyperlinks
            .push(Hyperlink {
                range: CellRange::parse_a1("A3:B3").unwrap(),
                external_target: None,
                location: Some("Sheet1!A1".into()),
                tooltip: None,
                display: None,
            });
        let before = wb.sheets[0].used_range();

        let op = Op::InsertRows {
            sheet: SheetId(0),
            at: 1,
            count: 2,
        };
        let inv = apply(&mut wb, &op).unwrap();
        assert_eq!(
            wb.value(SheetId(0), r("A5")),
            CellValue::Text { value: "x".into() }
        );
        assert_eq!(wb.sheets[0].hyperlinks[0].range.to_a1(), "A5:B5");

        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.sheets[0].used_range(), before);
        assert_eq!(
            wb.value(SheetId(0), r("A3")),
            CellValue::Text { value: "x".into() }
        );
        assert_eq!(wb.sheets[0].hyperlinks[0].range.to_a1(), "A3:B3");
    }

    #[test]
    fn remove_hyperlink_removes_all_duplicates_and_undo_restores_their_order() {
        let mut wb = wb_one_sheet();
        let link = |address: &str| Hyperlink {
            range: CellRange::parse_a1(address).unwrap(),
            external_target: None,
            location: Some("Data!D1".into()),
            tooltip: None,
            display: None,
        };
        wb.sheets[0].hyperlinks = vec![link("A1"), link("D4"), link("A1")];
        let before = wb.clone();
        let inverse = apply_in_place(
            &mut wb,
            &Op::RemoveHyperlink {
                sheet: SheetId(0),
                range: CellRange::parse_a1("$A$1").unwrap(),
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].hyperlinks, vec![link("D4")]);
        for op in inverse.0 {
            apply_in_place(&mut wb, &op).unwrap();
        }
        assert_eq!(wb, before);
    }

    #[test]
    fn structural_edits_and_renames_rewrite_hyperlink_destinations() {
        let mut wb = wb_one_sheet();
        wb.sheets[0].hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("A1").unwrap(),
            external_target: Some("https://example.com/report".into()),
            location: Some("Target!A3".into()),
            tooltip: None,
            display: Some("Jump".into()),
        });
        wb.sheets.push(Sheet::new("Target"));
        wb.sheets[1].hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("A1:A4").unwrap(),
            external_target: Some("https://example.com".into()),
            location: None,
            tooltip: None,
            display: None,
        });
        let before = wb.clone();

        let inverse = apply(
            &mut wb,
            &Op::InsertRows {
                sheet: SheetId(1),
                at: 1,
                count: 2,
            },
        )
        .unwrap();
        assert_eq!(
            wb.sheets[0].hyperlinks[0].location.as_deref(),
            Some("Target!A5")
        );
        assert_eq!(
            wb.sheets[0].hyperlinks[0].external_target.as_deref(),
            Some("https://example.com/report")
        );
        assert_eq!(wb.sheets[1].hyperlinks[0].range.to_a1(), "A1:A6");
        for operation in &inverse.0 {
            apply(&mut wb, operation).unwrap();
        }
        assert_eq!(wb, before);

        let inverse = apply(
            &mut wb,
            &Op::DeleteRows {
                sheet: SheetId(1),
                at: 0,
                count: 1,
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[1].hyperlinks[0].range.to_a1(), "A1:A3");
        assert_eq!(
            wb.sheets[0].hyperlinks[0].location.as_deref(),
            Some("Target!A2")
        );
        for operation in &inverse.0 {
            apply(&mut wb, operation).unwrap();
        }
        assert_eq!(wb, before);

        let inverse = apply(
            &mut wb,
            &Op::RenameSheet {
                sheet: SheetId(1),
                name: "New Target".into(),
            },
        )
        .unwrap();
        assert_eq!(
            wb.sheets[0].hyperlinks[0].location.as_deref(),
            Some("'New Target'!A3")
        );
        assert_eq!(
            wb.sheets[0].hyperlinks[0].external_target.as_deref(),
            Some("https://example.com/report")
        );
        for operation in &inverse.0 {
            apply(&mut wb, operation).unwrap();
        }
        assert_eq!(wb, before);
    }

    #[test]
    fn rename_rewrites_hash_prefixed_hyperlink_locations() {
        let mut wb = wb_one_sheet();
        let link = |range: &str, location: &str| Hyperlink {
            range: CellRange::parse_a1(range).unwrap(),
            external_target: None,
            location: Some(location.into()),
            tooltip: None,
            display: None,
        };
        wb.sheets[0].hyperlinks.extend([
            link("A1", "#Target!A1"),
            link("A2", "Target!A2"),
            link("A3", "#'Target'!A3"),
            link("A4", "#MyRange"),
            link("A5", "#Other!A1"),
        ]);
        wb.sheets[0].hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("A6").unwrap(),
            external_target: Some("https://example.com/report".into()),
            location: Some("#Target!A6".into()),
            tooltip: None,
            display: None,
        });
        wb.sheets.push(Sheet::new("Target"));
        let before = wb.clone();

        let inverse = apply(
            &mut wb,
            &Op::RenameSheet {
                sheet: SheetId(1),
                name: "My Sheet".into(),
            },
        )
        .unwrap();

        let locations: Vec<Option<&str>> = wb.sheets[0]
            .hyperlinks
            .iter()
            .map(|hyperlink| hyperlink.location.as_deref())
            .collect();
        assert_eq!(
            locations,
            vec![
                Some("#'My Sheet'!A1"),
                Some("'My Sheet'!A2"),
                Some("#'My Sheet'!A3"),
                Some("#MyRange"),
                Some("#Other!A1"),
                Some("#'My Sheet'!A6"),
            ]
        );
        assert_eq!(
            wb.sheets[0].hyperlinks[5].external_target.as_deref(),
            Some("https://example.com/report")
        );

        for operation in &inverse.0 {
            apply(&mut wb, operation).unwrap();
        }
        assert_eq!(wb, before);
    }

    #[test]
    fn column_style_runs_survive_structural_undo_and_redo() {
        let runs = vec![
            ColStyle {
                first: 0,
                last: 6,
                xf: 1,
            },
            ColStyle {
                first: 2,
                last: 3,
                xf: 2,
            },
            ColStyle {
                first: 3,
                last: 7,
                xf: 3,
            },
            ColStyle {
                first: MAX_COLS - 1,
                last: MAX_COLS - 1,
                xf: 4,
            },
        ];
        for op in [
            Op::DeleteCols {
                sheet: SheetId(0),
                at: 2,
                count: 2,
            },
            Op::DeleteCols {
                sheet: SheetId(0),
                at: 0,
                count: 8,
            },
            Op::InsertCols {
                sheet: SheetId(0),
                at: 3,
                count: 2,
            },
            Op::RemoveSheet { index: 0 },
        ] {
            let mut workbook = wb_one_sheet();
            workbook.sheets[0].col_styles = runs.clone();
            let before = workbook.clone();
            let inverse = apply_ops(&mut workbook, &[op]).unwrap();
            let edited = workbook.clone();
            let redo = apply_ops(&mut workbook, &inverse).unwrap();
            assert_eq!(workbook, before);
            let undo = apply_ops(&mut workbook, &redo).unwrap();
            assert_eq!(workbook, edited);
            apply_ops(&mut workbook, &undo).unwrap();
            assert_eq!(workbook, before);
        }
    }

    #[test]
    fn column_styles_follow_the_columns_they_format() {
        let mut wb = wb_one_sheet();
        let runs = vec![
            ColStyle {
                first: 0,
                last: 1,
                xf: 7,
            },
            ColStyle {
                first: 3,
                last: 5,
                xf: 9,
            },
        ];
        wb.sheet_mut(SheetId(0)).unwrap().col_styles = runs.clone();
        apply(
            &mut wb,
            &Op::InsertCols {
                sheet: SheetId(0),
                at: 1,
                count: 2,
            },
        )
        .unwrap();
        assert_eq!(
            wb.sheets[0].col_styles,
            vec![
                ColStyle {
                    first: 0,
                    last: 3,
                    xf: 7
                },
                ColStyle {
                    first: 5,
                    last: 7,
                    xf: 9
                },
            ],
            "the run the insert splits widens; the one to its right shifts"
        );
        assert_eq!(wb.sheets[0].col_style(3), Some(7));
        assert_eq!(wb.sheets[0].col_style(4), None);

        apply(
            &mut wb,
            &Op::DeleteCols {
                sheet: SheetId(0),
                at: 1,
                count: 2,
            },
        )
        .unwrap();
        assert_eq!(
            wb.sheets[0].col_styles, runs,
            "the delete undoes the insert"
        );

        apply(
            &mut wb,
            &Op::DeleteCols {
                sheet: SheetId(0),
                at: 3,
                count: 3,
            },
        )
        .unwrap();
        assert_eq!(
            wb.sheets[0].col_styles,
            vec![ColStyle {
                first: 0,
                last: 1,
                xf: 7
            }],
            "a run the delete consumes whole is dropped, as excel drops it"
        );

        wb.sheet_mut(SheetId(0)).unwrap().col_styles.push(ColStyle {
            first: 2,
            last: 9,
            xf: 9,
        });
        apply(
            &mut wb,
            &Op::DeleteCols {
                sheet: SheetId(0),
                at: 0,
                count: 4,
            },
        )
        .unwrap();
        assert_eq!(
            wb.sheets[0].col_styles,
            vec![ColStyle {
                first: 0,
                last: 5,
                xf: 9
            }],
            "a clipped run keeps the columns that survive"
        );
    }

    #[test]
    fn add_and_remove_sheet_round_trip() {
        let mut wb = wb_one_sheet();
        let op = Op::AddSheet {
            index: 1,
            name: "New".into(),
        };
        let inv = apply(&mut wb, &op).unwrap();
        assert_eq!(wb.sheets.len(), 2);
        assert_eq!(wb.sheets[1].name, "New");
        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.sheets.len(), 1);
    }

    #[test]
    fn remove_populated_sheet_round_trips_contents() {
        let mut wb = wb_one_sheet();
        wb.sheets.push(Sheet::new("Two"));
        let s = wb.sheet_mut(SheetId(1)).unwrap();
        s.set_cell(
            r("A1"),
            Cell {
                value: CellValue::Number { value: 9.0 },
                ..Default::default()
            },
        );
        s.merges.push(CellRange::parse_a1("A1:B1").unwrap());
        s.col_widths.insert(0, 12.0);
        s.freeze_pane = Some(FreezePane::new(1, 1, r("C4")));
        let before = wb.sheets[1].clone();

        let inv = apply(&mut wb, &Op::RemoveSheet { index: 1 }).unwrap();
        assert_eq!(wb.sheets.len(), 1);
        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.sheets.len(), 2);
        assert_eq!(
            wb.value(SheetId(1), r("A1")),
            CellValue::Number { value: 9.0 }
        );
        assert_eq!(wb.sheets[1].merges, before.merges);
        assert_eq!(wb.sheets[1].col_widths, before.col_widths);
        assert_eq!(wb.sheets[1].freeze_pane, before.freeze_pane);
    }

    #[test]
    fn merge_and_col_width_round_trip() {
        let mut wb = wb_one_sheet();
        let range = CellRange::parse_a1("A1:B2").unwrap();
        let inv = apply(
            &mut wb,
            &Op::MergeCells {
                sheet: SheetId(0),
                range,
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].merges, vec![range]);
        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert!(wb.sheets[0].merges.is_empty());

        let inv = apply(
            &mut wb,
            &Op::SetColWidth {
                sheet: SheetId(0),
                col: 3,
                width: Some(20.0),
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].col_widths.get(&3), Some(&20.0));
        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.sheets[0].col_widths.get(&3), None);
    }

    #[test]
    fn rename_sheet_round_trip() {
        let mut wb = wb_one_sheet();
        wb.sheets.push(Sheet::new("Formula"));
        wb.sheet_mut(SheetId(1)).unwrap().set_cell(
            r("A1"),
            Cell {
                formula: Some("Future!A1+Sheet1!A1".into()),
                ..Cell::default()
            },
        );
        let inv = apply(
            &mut wb,
            &Op::RenameSheet {
                sheet: SheetId(0),
                name: "Future".into(),
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].name, "Future");
        assert_eq!(wb.formula(SheetId(1), r("A1")), Some("Future!A1+Future!A1"));
        let redo = apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb.sheets[0].name, "Sheet1");
        assert_eq!(wb.formula(SheetId(1), r("A1")), Some("Future!A1+Sheet1!A1"));
        apply_ops(&mut wb, &redo).unwrap();
        assert_eq!(wb.sheets[0].name, "Future");
        assert_eq!(wb.formula(SheetId(1), r("A1")), Some("Future!A1+Future!A1"));
    }

    #[test]
    fn rename_undo_does_not_change_destination_only_references() {
        let mut wb = wb_one_sheet();
        wb.sheets.push(Sheet::new("Formula"));
        wb.sheet_mut(SheetId(1)).unwrap().set_cell(
            r("A1"),
            Cell {
                formula: Some("Future!A1".into()),
                ..Cell::default()
            },
        );
        wb.sheet_mut(SheetId(1)).unwrap().set_cell(
            r("A2"),
            Cell {
                formula: Some("Sheet1!A1".into()),
                ..Cell::default()
            },
        );

        let inverse = apply(
            &mut wb,
            &Op::RenameSheet {
                sheet: SheetId(0),
                name: "Future".into(),
            },
        )
        .unwrap();
        match &inverse.0[0] {
            Op::RestoreSheet { formulas, .. } => assert_eq!(formulas.len(), 1),
            other => panic!("expected restore sheet inverse, got {other:?}"),
        }
        apply_ops(&mut wb, &inverse.0).unwrap();

        assert_eq!(wb.formula(SheetId(1), r("A1")), Some("Future!A1"));
        assert_eq!(wb.formula(SheetId(1), r("A2")), Some("Sheet1!A1"));
    }

    #[test]
    fn rename_history_stays_constant_across_undo_redo() {
        let mut wb = wb_one_sheet();
        wb.sheets.push(Sheet::new("Formula"));
        wb.sheet_mut(SheetId(1)).unwrap().set_cell(
            r("A1"),
            Cell {
                formula: Some("Sheet1!A1".into()),
                ..Cell::default()
            },
        );
        let inverse = apply(
            &mut wb,
            &Op::RenameSheet {
                sheet: SheetId(0),
                name: "Future".into(),
            },
        )
        .unwrap();
        assert_eq!(inverse.0.len(), 1);
        let json = serde_json::to_string(&inverse.0).unwrap();
        let decoded: Vec<Op> = serde_json::from_str(&json).unwrap();
        assert_eq!(decoded, inverse.0);
        let redo = apply_ops(&mut wb, &inverse.0).unwrap();
        assert_eq!(redo.len(), 1);
        let undo = apply_ops(&mut wb, &redo).unwrap();
        assert_eq!(undo.len(), 1);
    }

    #[test]
    fn missing_sheet_errors() {
        let mut wb = Workbook::default();
        let err = apply(
            &mut wb,
            &Op::SetCell {
                sheet: SheetId(3),
                at: r("A1"),
                cell: num(1.0),
            },
        );
        assert_eq!(err.unwrap_err(), OpError::SheetNotFound(SheetId(3)));
    }

    #[test]
    fn apply_ops_is_atomic_when_formula_rewriting_fails() {
        let mut wb = wb_one_sheet();
        let before = wb.clone();
        let error = apply_ops(
            &mut wb,
            &[
                Op::SetCell {
                    sheet: SheetId(0),
                    at: r("A1"),
                    cell: CellState {
                        formula: Some("SUM(".into()),
                        ..CellState::default()
                    },
                },
                Op::InsertRows {
                    sheet: SheetId(0),
                    at: 0,
                    count: 1,
                },
            ],
        )
        .unwrap_err();
        assert!(matches!(error, OpError::FormulaNotRewritable { .. }));
        assert_eq!(wb, before);
    }

    #[test]
    fn structural_edits_capture_defined_name_rewrites_for_undo() {
        let mut wb = wb_one_sheet();
        wb.defined_names.push(DefinedName {
            name: "Input".into(),
            formula: "Sheet1!$A$2".into(),
            local_sheet: None,
            hidden: false,
        });
        let before = wb.clone();

        let inverse = apply(
            &mut wb,
            &Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        )
        .unwrap();

        assert_eq!(wb.defined_names[0].formula, "Sheet1!$A$3");
        assert!(
            inverse
                .0
                .iter()
                .any(|op| matches!(op, Op::SetDefinedNames { .. }))
        );
        apply_ops(&mut wb, &inverse.0).unwrap();
        assert_eq!(wb, before);
    }

    #[test]
    fn structural_edits_refuse_workbook_names_with_unqualified_references() {
        let mut wb = wb_one_sheet();
        wb.sheets.push(Sheet::new("Other"));
        wb.defined_names.push(DefinedName {
            name: "Input".into(),
            formula: "$A$2".into(),
            local_sheet: None,
            hidden: false,
        });
        let before = wb.clone();

        let error = apply(
            &mut wb,
            &Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        )
        .unwrap_err();

        assert!(matches!(error, OpError::DefinedNameNotRewritable { .. }));
        assert_eq!(wb, before);
    }

    /// A dynamic-range name is ordinary, and the row it names sits nowhere
    /// near the edit; the whole sheet used to be frozen by it all the same.
    #[test]
    fn a_name_the_formula_parser_cannot_read_still_accepts_the_edit() {
        let mut wb = wb_one_sheet();
        wb.sheets.push(Sheet::new("Other"));
        wb.defined_names.push(DefinedName {
            name: "Region".into(),
            formula: "Sheet1!$A$1:INDEX(Sheet1!$A:$A,COUNTA(Sheet1!$A:$A))".into(),
            local_sheet: None,
            hidden: false,
        });
        let before = wb.clone();

        apply(
            &mut wb,
            &Op::InsertRows {
                sheet: SheetId(0),
                at: 9998,
                count: 1,
            },
        )
        .unwrap();
        assert_eq!(wb.defined_names, before.defined_names);

        let inverse = apply(
            &mut wb,
            &Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        )
        .unwrap();
        assert_eq!(
            wb.defined_names[0].formula,
            "Sheet1!$A$2:INDEX(Sheet1!$A:$A,COUNTA(Sheet1!$A:$A))"
        );
        apply_ops(&mut wb, &inverse.0).unwrap();
        assert_eq!(wb.defined_names, before.defined_names);
    }

    #[test]
    fn refused_defined_name_rewrites_leave_the_workbook_unchanged() {
        let mut wb = wb_one_sheet();
        wb.sheets[0].set_cell(
            r("B1"),
            Cell {
                formula: Some("A2".into()),
                ..Cell::default()
            },
        );
        wb.defined_names.push(DefinedName {
            name: "Dynamic".into(),
            formula: "SUM(Table1[Amount])".into(),
            local_sheet: Some(SheetId(0)),
            hidden: false,
        });
        let before = wb.clone();

        let error = apply(
            &mut wb,
            &Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        )
        .unwrap_err();

        assert!(matches!(error, OpError::DefinedNameNotRewritable { .. }));
        assert_eq!(wb, before);
    }

    #[test]
    fn setting_a_chart_anchor_inverts_to_the_previous_one() {
        use xlsx_model::{AnchorCell, AnchorEditAs, ChartAnchor, SheetChart};

        let anchor = |col, row| ChartAnchor::TwoCell {
            from: AnchorCell {
                col,
                row,
                ..AnchorCell::default()
            },
            to: AnchorCell {
                col: col + 4,
                row: row + 8,
                ..AnchorCell::default()
            },
            edit_as: AnchorEditAs::TwoCell,
        };
        let mut wb = wb_one_sheet();
        wb.sheets[0].charts.push(SheetChart {
            part: "xl/charts/chart1.xml".into(),
            drawing: "xl/drawings/drawing1.xml".into(),
            anchor_index: 0,
            anchor: anchor(0, 0),
            refs: Vec::new(),
        });

        let op = Op::SetChartAnchor {
            sheet: SheetId(0),
            frame: "xl/drawings/drawing1.xml#0".into(),
            part: "xl/charts/chart1.xml".into(),
            from: anchor(0, 0),
            to: anchor(3, 5),
        };
        let inverse = apply(&mut wb, &op).unwrap();
        assert_eq!(wb.sheets[0].charts[0].anchor, anchor(3, 5));
        for op in &inverse.0 {
            apply(&mut wb, op).unwrap();
        }
        assert_eq!(wb.sheets[0].charts[0].anchor, anchor(0, 0));

        let error = apply(
            &mut wb,
            &Op::SetChartAnchor {
                sheet: SheetId(0),
                frame: "xl/drawings/drawing9.xml#0".into(),
                part: "xl/charts/chart1.xml".into(),
                from: anchor(0, 0),
                to: anchor(1, 1),
            },
        )
        .unwrap_err();
        assert!(matches!(error, OpError::ChartNotFound { .. }));

        // the frame is there but holds neither the anchor nor the part the op
        // recorded: the ordinal now names something else, so it is refused.
        for (part, from) in [
            ("xl/charts/chart1.xml", anchor(7, 7)),
            ("xl/charts/chart9.xml", anchor(0, 0)),
        ] {
            let error = apply(
                &mut wb,
                &Op::SetChartAnchor {
                    sheet: SheetId(0),
                    frame: "xl/drawings/drawing1.xml#0".into(),
                    part: part.into(),
                    from,
                    to: anchor(1, 1),
                },
            )
            .unwrap_err();
            assert!(matches!(error, OpError::ChartFrameShifted { .. }));
            assert_eq!(wb.sheets[0].charts[0].anchor, anchor(0, 0));
        }
    }

    fn filled(at: &str, value: f64) -> (CellRef, Cell) {
        (
            r(at),
            Cell {
                value: CellValue::Number { value },
                ..Cell::default()
            },
        )
    }

    /// two cells land on one target; today's clear-then-reinsert order made
    /// the largest source address win.
    #[test]
    fn remap_cells_contested_target_keeps_the_largest_source() {
        let mut sheet = Sheet::new("S");
        let (a1, c1) = filled("A1", 1.0);
        let (_, c2) = filled("A6", 2.0);
        sheet.set_cell(a1, c1.clone());
        sheet.set_cell(r("A6"), c2.clone());
        let dropped = sheet.remap_cells(|at| match at.row {
            0 => Some(CellRef::new(5, 0)),
            5 => Some(CellRef::new(5, 0)),
            _ => Some(at),
        });
        assert!(dropped.is_empty());
        assert_eq!(sheet.cell(CellRef::new(5, 0)), Some(&c2));
        assert_eq!(sheet.iter_cells().count(), 1);
    }

    /// the same contest when every touched cell moves, so the split-off
    /// fast path runs instead of the full-rebuild fallback.
    #[test]
    fn remap_cells_contested_target_resolves_on_the_fast_path_too() {
        let mut sheet = Sheet::new("S");
        let (_, c5) = filled("A5", 5.0);
        let (_, c7) = filled("A7", 7.0);
        sheet.set_cell(r("A5"), c5);
        sheet.set_cell(r("A7"), c7.clone());
        let dropped = sheet.remap_cells(|at| {
            if at.row >= 4 {
                Some(CellRef::new(6, at.col))
            } else {
                Some(at)
            }
        });
        assert!(dropped.is_empty());
        assert_eq!(sheet.cell(CellRef::new(6, 0)), Some(&c7));
        assert_eq!(sheet.iter_cells().count(), 1);
    }

    /// cells that keep their address stay put, refused ones come back in
    /// address order, and survivors move down past the refused span.
    #[test]
    fn remap_cells_keeps_stays_and_reports_drops_in_address_order() {
        let mut sheet = Sheet::new("S");
        for (at, value) in [("A1", 1.0), ("A2", 2.0), ("A3", 3.0), ("B7", 7.0)] {
            let (at, cell) = filled(at, value);
            sheet.set_cell(at, cell);
        }
        let dropped = sheet.remap_cells(|at| {
            if at.row == 1 {
                None
            } else if at.row == 0 {
                Some(at)
            } else {
                Some(CellRef::new(at.row - 1, at.col))
            }
        });
        assert_eq!(dropped.len(), 1);
        assert_eq!(dropped[0].0, r("A2"));
        assert_eq!(dropped[0].1.value, CellValue::Number { value: 2.0 });
        assert_eq!(
            sheet.cell(r("A1")).unwrap().value,
            CellValue::Number { value: 1.0 }
        );
        assert_eq!(
            sheet.cell(r("A2")).unwrap().value,
            CellValue::Number { value: 3.0 }
        );
        assert_eq!(
            sheet.cell(r("B6")).unwrap().value,
            CellValue::Number { value: 7.0 }
        );
        assert_eq!(sheet.iter_cells().count(), 3);
    }

    /// rows and columns shift through the new storage path and invert back to
    /// exactly the workbook they started from.
    #[test]
    fn structural_shifts_round_trip_exactly_on_both_axes() {
        let mut wb = wb_one_sheet();
        wb.sheets.push(Sheet::new("Data"));
        for row in 1..=12 {
            for col in ["A", "C", "F"] {
                let at = format!("{col}{row}");
                let (cell_ref, cell) = filled(&at, f64::from(row));
                wb.sheet_mut(SheetId(0)).unwrap().set_cell(cell_ref, cell);
            }
            wb.sheet_mut(SheetId(1)).unwrap().set_cell(
                r(&format!("B{row}")),
                Cell {
                    formula: Some("Sheet1!A1+Sheet1!$E$9".into()),
                    ..Cell::default()
                },
            );
        }
        let before = wb.clone();

        let edits = [
            Op::InsertRows {
                sheet: SheetId(0),
                at: 4,
                count: 3,
            },
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 1,
                count: 2,
            },
            Op::InsertCols {
                sheet: SheetId(0),
                at: 1,
                count: 2,
            },
            Op::DeleteCols {
                sheet: SheetId(0),
                at: 3,
                count: 1,
            },
        ];
        for op in &edits {
            let inverse = apply(&mut wb, op).unwrap();
            apply_ops(&mut wb, &inverse.0).unwrap();
            assert_eq!(wb, before, "{op:?} must invert exactly");
        }
    }

    /// a deletion whose span holds contents must report them for the inverse,
    /// whatever sits above or below the span.
    #[test]
    fn deleting_rows_reports_every_dropped_cell_for_undo() {
        let mut wb = wb_one_sheet();
        for (at, value) in [("A1", 1.0), ("A3", 3.0), ("A4", 4.0), ("A6", 6.0)] {
            let (at, cell) = filled(at, value);
            wb.sheets[0].set_cell(at, cell);
        }
        let inverse = apply(
            &mut wb,
            &Op::DeleteRows {
                sheet: SheetId(0),
                at: 2,
                count: 2,
            },
        )
        .unwrap();
        assert_eq!(
            wb.value(SheetId(0), r("A4")),
            CellValue::Number { value: 6.0 }
        );
        assert!(wb.sheet(SheetId(0)).unwrap().cell(r("A3")).is_none());
        let restores = inverse
            .0
            .iter()
            .filter(|op| matches!(op, Op::SetCell { .. }))
            .count();
        assert_eq!(restores, 2, "rows 3 and 4 must come back");
        apply_ops(&mut wb, &inverse.0).unwrap();
        assert_eq!(
            wb.value(SheetId(0), r("A3")),
            CellValue::Number { value: 3.0 }
        );
        assert_eq!(
            wb.value(SheetId(0), r("A4")),
            CellValue::Number { value: 4.0 }
        );
    }

    #[test]
    fn number_format_matching_the_base_xf_is_stored_on_an_unstyled_cell() {
        let mut wb = wb_one_sheet();
        wb.styles.cell_xfs.push(xlsx_model::Xf {
            num_fmt_id: Some(10),
            ..Default::default()
        });
        wb.sheets[0].set_cell(r("A1"), num(0.5).into());
        apply(
            &mut wb,
            &Op::SetRangeNumberFormat {
                sheet: SheetId(0),
                range: CellRange::new(r("A1"), r("A1")),
                format: crate::NumberFormatMutation::Percent,
            },
        )
        .unwrap();
        let style = wb.sheets[0].cell(r("A1")).unwrap().style;
        assert_eq!(
            wb.styles
                .cell_format(Some(style.unwrap_or(0)))
                .number_format,
            xlsx_model::NumberFormat::Builtin { id: 10 }
        );
    }

    #[test]
    fn number_format_edit_matching_a_base_xf_differing_only_in_font_keeps_the_cell_unstyled() {
        let mut wb = wb_one_sheet();
        wb.styles.fonts.push(xlsx_model::Font {
            name: Some("Calibri".into()),
            size_pt: Some(11.0),
            ..Default::default()
        });
        wb.styles.cell_xfs.push(xlsx_model::Xf {
            font: Some(0),
            ..Default::default()
        });
        wb.sheets[0].set_cell(r("A1"), num(0.5).into());
        let inverse = apply(
            &mut wb,
            &Op::SetRangeNumberFormat {
                sheet: SheetId(0),
                range: CellRange::new(r("A1"), r("A2")),
                format: crate::NumberFormatMutation::Automatic,
            },
        )
        .unwrap();
        assert!(inverse.0.is_empty());
        assert_eq!(wb.sheets[0].cell(r("A1")).unwrap().style, None);
        assert!(wb.sheets[0].cell(r("A2")).is_none());
    }
}

#[cfg(test)]
mod range_format_exhaustion_tests {
    use super::*;
    use crate::formatting::CapturedFormat;
    use xlsx_model::{CellFormat, CellRange, NumberFormat, SheetId};

    #[test]
    fn captured_format_differing_from_provenance_applies_its_format() {
        let mut wb = Workbook::default();
        wb.sheets.push(xlsx_model::Sheet::new("Sheet1"));
        wb.styles.cell_xfs.push(xlsx_model::Xf::default());
        let original = CellFormat {
            number_format: NumberFormat::Builtin { id: 1 },
            ..CellFormat::default()
        };
        let source_style = wb.styles.intern_derived_cell_format(0, &original).unwrap();
        let source =
            crate::formatting::PreservedCellFormat::from_style(&wb.styles, Some(source_style));
        let target = CellFormat {
            quote_prefix: true,
            ..CellFormat::default()
        };
        let target_style = wb.styles.intern_cell_format(&target).unwrap();
        wb.sheets[0].set_cell(
            xlsx_model::CellRef::new(0, 1),
            xlsx_model::Cell {
                style: target_style,
                ..Default::default()
            },
        );
        let applied = CellFormat {
            number_format: NumberFormat::Builtin { id: 2 },
            ..original
        };
        apply(
            &mut wb,
            &Op::ApplyRangeFormat {
                sheet: SheetId(0),
                range: CellRange::new(
                    xlsx_model::CellRef::new(0, 0),
                    xlsx_model::CellRef::new(0, 1),
                ),
                format: CapturedFormat {
                    rows: 1,
                    columns: 1,
                    formats: vec![applied.clone()],
                    kept: vec![crate::formatting::KeptFacets {
                        quote_prefix: true,
                        ..Default::default()
                    }],
                    source_styles: vec![Some(source_style)],
                    source_formats: vec![source],
                },
            },
        )
        .unwrap();
        for col in 0..2 {
            let style = wb.sheets[0]
                .cell(xlsx_model::CellRef::new(0, col))
                .unwrap()
                .style;
            let expected = CellFormat {
                quote_prefix: col == 1,
                ..applied.clone()
            };
            assert_eq!(wb.styles.cell_format(style), expected);
            assert_ne!(style, Some(source_style));
            assert_eq!(wb.styles.format_source(style.unwrap()), (None, Some(0)));
        }
    }

    #[test]
    fn mid_range_format_exhaustion_leaves_the_workbook_unchanged() {
        let mut styles = xlsx_model::Stylesheet::default();
        styles.num_fmts = (164..u16::MAX).map(|id| (id, format!("p{id}"))).collect();
        let mut wb = Workbook {
            styles,
            ..Workbook::default()
        };
        wb.sheets.push(xlsx_model::Sheet::new("Sheet1"));
        let sheet = SheetId(0);
        let before = wb.clone();

        let op = Op::ApplyRangeFormat {
            sheet,
            range: CellRange::new(
                xlsx_model::CellRef::new(0, 0),
                xlsx_model::CellRef::new(0, 1),
            ),
            format: CapturedFormat {
                source_styles: Vec::new(),
                source_formats: Vec::new(),
                rows: 1,
                columns: 2,
                formats: vec![
                    CellFormat {
                        number_format: NumberFormat::Custom {
                            pattern: "one".into(),
                        },
                        ..CellFormat::default()
                    },
                    CellFormat {
                        number_format: NumberFormat::Custom {
                            pattern: "two".into(),
                        },
                        ..CellFormat::default()
                    },
                ],
                kept: Vec::new(),
            },
        };

        assert!(matches!(
            apply_ops(&mut wb, &[op]),
            Err(OpError::NumFmtTableFull)
        ));
        assert_eq!(wb, before, "exhaustion must leave the workbook untouched");
    }

    /// `apply_ops` stages on its own clone, so only the single-op entry point
    /// proves `apply_range_formats` rolls its own pool writes back.
    #[test]
    fn mid_range_exhaustion_rolls_back_pools_without_an_outer_clone() {
        let mut styles = xlsx_model::Stylesheet::default();
        styles.num_fmts = (164..u16::MAX).map(|id| (id, format!("p{id}"))).collect();
        let mut wb = Workbook {
            styles,
            ..Workbook::default()
        };
        wb.sheets.push(xlsx_model::Sheet::new("Sheet1"));
        let before = wb.clone();

        let op = Op::ApplyRangeFormat {
            sheet: SheetId(0),
            range: CellRange::new(
                xlsx_model::CellRef::new(0, 0),
                xlsx_model::CellRef::new(0, 1),
            ),
            format: CapturedFormat {
                source_styles: Vec::new(),
                source_formats: Vec::new(),
                rows: 1,
                columns: 2,
                formats: vec![
                    CellFormat {
                        number_format: NumberFormat::Custom {
                            pattern: "one".into(),
                        },
                        ..CellFormat::default()
                    },
                    CellFormat {
                        number_format: NumberFormat::Custom {
                            pattern: "two".into(),
                        },
                        ..CellFormat::default()
                    },
                ],
                kept: Vec::new(),
            },
        };

        assert!(matches!(apply(&mut wb, &op), Err(OpError::NumFmtTableFull)));
        assert_eq!(
            wb, before,
            "a failed range format must not leave interned entries behind"
        );
    }
}

#[cfg(test)]
mod merge_c1300_tests {
    use super::*;
    use crate::{CapturedFormat, NumberFormatMutation, PreservedCellFormat, StylePatch};
    use xlsx_model::{CellFormat, Color, Fill, NumberFormat, Xf};

    fn workbook() -> Workbook {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Sheet1"));
        wb.styles.cell_xfs.push(Xf::default());
        wb
    }

    #[test]
    fn captured_plain_under_a_styled_row_stores_xf_zero() {
        let mut wb = workbook();
        let row_format = CellFormat {
            fill: Fill::Solid(Color::Rgb("#123456".into())),
            ..CellFormat::default()
        };
        let row_style = wb.styles.intern_cell_format(&row_format).unwrap().unwrap();
        wb.sheets[0].row_styles.insert(0, row_style);
        let source = PreservedCellFormat::from_style(&wb.styles, None);
        assert_eq!(source.format, CellFormat::default());
        assert_eq!(source.source_xf, None);
        assert_eq!(source.derived_from, None);
        let inverse = apply(
            &mut wb,
            &Op::ApplyRangeFormat {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:B2").unwrap(),
                format: CapturedFormat {
                    rows: 1,
                    columns: 1,
                    formats: vec![CellFormat::default()],
                    kept: Vec::new(),
                    source_styles: vec![None],
                    source_formats: vec![source],
                },
            },
        )
        .unwrap();
        for col in 0..2 {
            let at = CellRef::new(0, col);
            assert_eq!(wb.sheets[0].cell(at).map(|cell| cell.style), Some(Some(0)));
            let displayed = wb.sheets[0].effective_style(at);
            assert_eq!(wb.styles.cell_format(displayed), CellFormat::default());
            assert!(wb.sheets[0].cell(CellRef::new(1, col)).is_none());
        }
        assert_eq!(inverse.0.len(), 2);
        apply_ops(&mut wb, &inverse.0).unwrap();
        assert_eq!(wb.sheets[0].iter_cells().count(), 0);
        assert_eq!(
            wb.sheets[0].effective_style(CellRef::new(0, 0)),
            Some(row_style)
        );
    }

    #[test]
    fn partial_noop_over_a_nonplain_xf_zero_stays_sparse() {
        let mut wb = workbook();
        wb.styles.cell_xfs[0].num_fmt_id = Some(10);
        let before = wb.clone();
        let inverse = apply(
            &mut wb,
            &Op::SetRangeNumberFormat {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:C3").unwrap(),
                format: NumberFormatMutation::Percent,
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].iter_cells().count(), 0);
        assert!(inverse.0.is_empty());
        assert_eq!(wb, before);
    }

    #[test]
    fn partial_edit_on_a_missing_cell_derives_from_its_row() {
        let mut wb = workbook();
        let row_format = CellFormat {
            fill: Fill::Solid(Color::Rgb("#123456".into())),
            number_format: NumberFormat::Builtin { id: 10 },
            ..CellFormat::default()
        };
        let row_style = wb.styles.intern_cell_format(&row_format).unwrap().unwrap();
        wb.sheets[0].row_styles.insert(0, row_style);
        let at = CellRef::new(0, 0);
        let inverse = apply(
            &mut wb,
            &Op::PatchRangeStyle {
                sheet: SheetId(0),
                range: CellRange::new(at, at),
                patch: StylePatch {
                    bold: Some(true),
                    ..StylePatch::default()
                },
            },
        )
        .unwrap();
        let style = wb.sheets[0].cell(at).unwrap().style.unwrap();
        let actual = wb.styles.cell_format(Some(style));
        assert_eq!(actual.fill, row_format.fill);
        assert_eq!(actual.number_format, row_format.number_format);
        assert!(actual.font.bold);
        assert_eq!(wb.styles.format_source(style), (None, Some(row_style)));
        assert_eq!(inverse.0.len(), 1);
        apply_ops(&mut wb, &inverse.0).unwrap();
        assert!(wb.sheets[0].cell(at).is_none());
        assert_eq!(wb.sheets[0].effective_style(at), Some(row_style));
    }
}
