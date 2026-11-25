use std::collections::HashSet;

use xlsx_calc::lexer::rewrite_table_column;
use xlsx_model::workbook::{MAX_DEFINED_NAME_BYTES, defined_names_bytes};
use xlsx_model::{CellRange, CellRef, CellValue, SheetId, Table, Workbook};

use crate::{CellState, Op, OpError};

pub fn table_header_at(tables: &[Table], sheet: SheetId, at: CellRef) -> Option<(usize, usize)> {
    tables.iter().enumerate().find_map(|(index, table)| {
        (table.sheet == sheet
            && table.header_rows != 0
            && at.row == table.range.start.row
            && (table.range.start.col..=table.range.end.col).contains(&at.col))
        .then(|| (index, (at.col - table.range.start.col) as usize))
    })
}

pub fn normalize_table_header(
    workbook: &Workbook,
    sheet: SheetId,
    at: CellRef,
    cell: &CellState,
) -> Result<Option<(usize, usize, String)>, OpError> {
    if workbook.tables.is_empty() {
        return Ok(None);
    }
    let old = workbook.sheet(sheet).and_then(|sheet| sheet.cell(at));
    if cell.formula.as_deref() == old.and_then(|cell| cell.formula.as_deref())
        && (cell.formula.is_some()
            || &cell.value == old.map_or(&CellValue::Empty, |cell| &cell.value))
    {
        return Ok(None);
    }
    let tables = &workbook.tables;
    let Some((table, column)) = table_header_at(tables, sheet, at) else {
        return Ok(None);
    };
    if cell.formula.is_some() {
        return Err(OpError::TableHeaderFormulaUnsupported);
    }
    let columns = &tables[table].columns;
    if column >= columns.len() {
        return Err(OpError::MissingTableColumn { sheet, cell: at });
    }
    let base = match &cell.value {
        CellValue::Empty => String::new(),
        CellValue::Text { value } => value.clone(),
        CellValue::Number { value } => value.to_string(),
        CellValue::Bool { value: true } => "TRUE".to_owned(),
        CellValue::Bool { value: false } => "FALSE".to_owned(),
        CellValue::Error { value } => value.as_str().to_owned(),
    };
    let taken = columns
        .iter()
        .enumerate()
        .filter(|(index, _)| *index != column)
        .map(|(_, name)| name.to_lowercase())
        .collect::<HashSet<_>>();
    let name = if base.is_empty() {
        (1..)
            .map(|suffix| format!("Column{suffix}"))
            .find(|name| !taken.contains(&name.to_lowercase()))
            .unwrap()
    } else if taken.contains(&base.to_lowercase()) {
        (2..)
            .map(|suffix| format!("{base}{suffix}"))
            .find(|name| !taken.contains(&name.to_lowercase()))
            .unwrap()
    } else {
        base
    };
    if name.chars().count() > xlsx_calc::eval::MAX_CELL_TEXT_CHARS {
        return Err(OpError::TableHeaderTextTooLong);
    }
    Ok(Some((table, column, name)))
}

#[doc(hidden)]
pub fn references_table_column(formula: &str, table: &str, local: bool, column: &str) -> bool {
    match rewrite_table_column(formula, table, local, column, column) {
        Ok(rewritten) => rewritten.is_some(),
        Err(_) => may_reference_column(formula, table, local, column),
    }
}

fn may_reference_column(formula: &str, table: &str, local: bool, column: &str) -> bool {
    let text = formula.to_lowercase();
    (local && text.contains('[') || text.contains(&format!("{}[", table.to_lowercase())))
        && (text.contains(&column.to_lowercase())
            || text
                .replace('\'', "")
                .contains(&column.to_lowercase().replace('\'', "")))
}

pub(crate) fn rename_table_column(
    workbook: &mut Workbook,
    table: usize,
    column: usize,
    at: CellRef,
    new: &str,
    inverse: &mut Vec<Op>,
) -> Result<(), OpError> {
    let table = &workbook.tables[table];
    let old = &table.columns[column];
    let changes_lookup = old.to_lowercase() != new.to_lowercase();
    let unsupported = |location| OpError::TableHeaderReferenceUnsupported {
        table: table.name.clone(),
        column: old.clone(),
        location,
    };
    let rewrite =
        |formula: &str, local: bool, location: &dyn Fn() -> String| match rewrite_table_column(
            formula,
            &table.name,
            local,
            old,
            new,
        ) {
            Ok(Some(text)) if !text.contains('\0') => Ok(Some(text)),
            Ok(None) => Ok(None),
            Err(_) if !may_reference_column(formula, &table.name, local, old) => Ok(None),
            _ => Err(unsupported(location())),
        };
    let mut edits = Vec::new();
    let mut restores = Vec::new();
    for (index, sheet) in workbook.sheets.iter().enumerate() {
        let id = SheetId(index as u32);
        if changes_lookup {
            for group in &sheet.conditional_formats {
                let local = id == table.sheet
                    && group.sqref.split_ascii_whitespace().any(|text| {
                        CellRange::parse_a1(text)
                            .is_ok_and(|range| table.range.contains(range.start))
                    });
                if group
                    .rules
                    .iter()
                    .flat_map(|rule| &rule.formulas)
                    .any(|text| references_table_column(text, &table.name, local, old))
                {
                    return Err(unsupported(format!(
                        "conditional formats on sheet '{}'",
                        sheet.name
                    )));
                }
            }
            for chart in &sheet.charts {
                if chart.refs.iter().any(|reference| {
                    references_table_column(&reference.formula, &table.name, false, old)
                }) {
                    return Err(unsupported(chart.part.clone()));
                }
            }
        }
        for (cell_ref, cell) in sheet.iter_cells() {
            if id == table.sheet && cell_ref == at {
                continue;
            }
            let Some(formula) = cell.formula.as_deref() else {
                continue;
            };
            let local = id == table.sheet && table.range.contains(cell_ref);
            let location = || format!("{}!{}", sheet.name, cell_ref.to_a1());
            let updated = rewrite(formula, local, &location)?;
            if updated.is_some() && table_header_at(&workbook.tables, id, cell_ref).is_some() {
                return Err(unsupported(location()));
            }
            if updated.is_some()
                || (changes_lookup && references_table_column(formula, &table.name, local, new))
            {
                restores.push(Op::SetCell {
                    sheet: id,
                    at: cell_ref,
                    cell: CellState::from(cell),
                });
            }
            if let Some(updated) = updated {
                edits.push((index, cell_ref, updated));
            }
        }
    }
    let mut names = None;
    let mut restore_names = false;
    for (index, name) in workbook.defined_names.iter().enumerate() {
        let location = || format!("defined name '{}'", name.name);
        if let Some(updated) = rewrite(&name.formula, false, &location)? {
            crate::validate_defined_name_formula(&updated).map_err(|_| unsupported(location()))?;
            names.get_or_insert_with(|| workbook.defined_names.clone())[index].formula = updated;
        } else if changes_lookup && references_table_column(&name.formula, &table.name, false, new)
        {
            restore_names = true;
        }
    }
    if let Some(names) = names {
        if defined_names_bytes(&names) > MAX_DEFINED_NAME_BYTES {
            return Err(OpError::TooManyDefinedNames);
        }
        inverse.push(Op::SetDefinedNames {
            defined_names: std::mem::replace(&mut workbook.defined_names, names),
        });
    } else if restore_names {
        inverse.push(Op::SetDefinedNames {
            defined_names: workbook.defined_names.clone(),
        });
    }
    inverse.extend(restores);
    for (sheet, at, formula) in edits {
        workbook.sheets[sheet].cell_mut(at).unwrap().formula = Some(formula);
    }
    Ok(())
}
