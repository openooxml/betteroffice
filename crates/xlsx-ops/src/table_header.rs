use std::collections::HashSet;

use xlsx_model::{CellRef, CellValue, SheetId, Table};

use crate::{CellState, OpError};

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
    tables: &[Table],
    sheet: SheetId,
    at: CellRef,
    cell: &CellState,
) -> Result<Option<(usize, usize, String)>, OpError> {
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
    Ok(Some((table, column, name)))
}
