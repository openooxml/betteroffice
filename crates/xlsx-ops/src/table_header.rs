use std::collections::HashSet;

use xlsx_calc::lexer::{may_reference_table_column, rewrite_table_column};
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

pub fn table_header_write_at(
    workbook: &Workbook,
    sheet: SheetId,
    at: CellRef,
    cell: &CellState,
) -> Option<(usize, usize)> {
    let header = table_header_at(&workbook.tables, sheet, at)?;
    let old = workbook.sheet(sheet).and_then(|sheet| sheet.cell(at));
    let same_content = cell.formula.as_deref() == old.and_then(|cell| cell.formula.as_deref())
        && (cell.formula.is_some()
            || &cell.value == old.map_or(&CellValue::Empty, |cell| &cell.value));
    let clear = cell.formula.is_none()
        && (matches!(&cell.value, CellValue::Empty)
            || matches!(&cell.value, CellValue::Text { value } if value.is_empty()));
    (!same_content || clear && cell.style == old.and_then(|cell| cell.style)).then_some(header)
}

pub fn normalize_table_header(
    workbook: &Workbook,
    sheet: SheetId,
    at: CellRef,
    cell: &CellState,
) -> Result<Option<(usize, usize, String)>, OpError> {
    let Some((table, column)) = table_header_write_at(workbook, sheet, at, cell) else {
        return Ok(None);
    };
    if cell.formula.is_some() {
        return Err(OpError::TableHeaderUnsupported(
            "formulas are not supported in Excel table headers".into(),
        ));
    }
    let columns = &workbook.tables[table].columns;
    if column >= columns.len() {
        return Err(OpError::TableHeaderUnsupported(format!(
            "no table column for sheet {}, {}",
            sheet.0,
            at.to_a1()
        )));
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
        return Err(OpError::TableHeaderUnsupported(
            "cell text exceeds Excel's length limit".into(),
        ));
    }
    Ok(Some((table, column, name)))
}

#[doc(hidden)]
pub fn references_table_column(
    formula: &str,
    table: &str,
    sheet: &str,
    local: bool,
    column: &str,
) -> bool {
    match rewrite_table_column(formula, table, sheet, local, column, column) {
        Ok(rewritten) => rewritten.is_some(),
        Err(_) => may_reference_table_column(formula, table, local, column),
    }
}

pub(crate) fn rename_table_column(
    workbook: &mut Workbook,
    index: usize,
    column: usize,
    at: CellRef,
    new: &str,
    cell: CellState,
) -> Result<Op, OpError> {
    let table = &workbook.tables[index];
    let old = &table.columns[column];
    let sheet = table.sheet;
    let name = table.name.clone();
    let mut formulas = Vec::new();
    let mut defined_names = Vec::new();
    if old.to_lowercase() != new.to_lowercase() {
        rewrite_table_formulas(
            workbook,
            index,
            column,
            at,
            new,
            &mut formulas,
            &mut defined_names,
        )?;
    }
    Ok(Op::RestoreTableColumn {
        sheet,
        at,
        table: name,
        caption: std::mem::replace(&mut workbook.tables[index].columns[column], new.to_owned()),
        cell,
        formulas,
        defined_names,
    })
}

type CellFormulaRestore = (SheetId, CellRef, (String, String));
type NameFormulaRestore = (xlsx_model::DefinedName, String);

fn rewrite_table_formulas(
    workbook: &mut Workbook,
    table: usize,
    column: usize,
    at: CellRef,
    new: &str,
    formulas: &mut Vec<CellFormulaRestore>,
    defined_names: &mut Vec<NameFormulaRestore>,
) -> Result<(), OpError> {
    let table = &workbook.tables[table];
    let old = &table.columns[column];
    let unsupported = |location: String| {
        OpError::TableHeaderUnsupported(format!(
            "cannot rename table {} column {old}: a reference in {location} \
             cannot be safely rewritten",
            table.name
        ))
    };
    let table_sheet = &workbook.sheets[table.sheet.0 as usize].name;
    let rewrite =
        |formula: &str, local: bool, location: &dyn Fn() -> String| match rewrite_table_column(
            formula,
            &table.name,
            table_sheet,
            local,
            old,
            new,
        ) {
            Ok(text) => Ok(text),
            Err(error) => Err(OpError::TableHeaderUnsupported(format!(
                "cannot rename table {} column {old}: a reference in {} \
                 cannot be safely rewritten: {error}",
                table.name,
                location()
            ))),
        };
    let mut edits = Vec::new();
    for (index, sheet) in workbook.sheets.iter().enumerate() {
        let id = SheetId(index as u32);
        for group in &sheet.conditional_formats {
            let local = id == table.sheet
                && group.sqref.split_ascii_whitespace().any(|text| {
                    CellRange::parse_a1(text).is_ok_and(|range| table.range.contains(range.start))
                });
            if group
                .rules
                .iter()
                .flat_map(|rule| &rule.formulas)
                .any(|text| references_table_column(text, &table.name, table_sheet, local, old))
            {
                return Err(unsupported(format!(
                    "conditional formats on sheet '{}'",
                    sheet.name
                )));
            }
        }
        for chart in &sheet.charts {
            if chart.refs.iter().any(|reference| {
                references_table_column(&reference.formula, &table.name, table_sheet, false, old)
            }) {
                return Err(unsupported(chart.part.clone()));
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
            if let Some(updated) = updated {
                edits.push((index, cell_ref, updated));
            }
        }
    }
    let mut names = Vec::new();
    let mut name_bytes = defined_names_bytes(&workbook.defined_names);
    for (index, name) in workbook.defined_names.iter().enumerate() {
        let location = || format!("defined name '{}'", name.name);
        if let Some(updated) = rewrite(&name.formula, false, &location)? {
            crate::validate_defined_name_formula(&updated).map_err(|_| unsupported(location()))?;
            name_bytes = name_bytes - name.formula.len() + updated.len();
            names.push((index, updated));
        }
    }
    if !names.is_empty() && name_bytes > MAX_DEFINED_NAME_BYTES {
        return Err(unsupported("defined names".into()));
    }
    for (index, formula) in names {
        let name = &mut workbook.defined_names[index];
        defined_names.push((name.clone(), formula.clone()));
        name.formula = formula;
    }
    for (sheet, at, formula) in edits {
        let stored = workbook.sheets[sheet].cell_mut(at).unwrap();
        let prior = stored.formula.replace(formula.clone()).unwrap();
        formulas.push((SheetId(sheet as u32), at, (prior, formula)));
    }
    Ok(())
}

pub(crate) fn restore_table_column(workbook: &mut Workbook, op: &Op) -> Result<Op, OpError> {
    let Op::RestoreTableColumn {
        sheet,
        at,
        table,
        caption,
        cell,
        formulas,
        defined_names,
    } = op
    else {
        unreachable!();
    };
    let missing = || {
        OpError::TableHeaderUnsupported(format!("table {table} has no header at {}", at.to_a1()))
    };
    let (index, column) = table_header_at(&workbook.tables, *sheet, *at)
        .filter(|(index, _)| workbook.tables[*index].name == *table)
        .ok_or_else(missing)?;
    let definition = &workbook.tables[index];
    let current = definition.columns.get(column).ok_or_else(missing)?;
    let sheet_name = &workbook.sheets[sheet.0 as usize].name;
    let restore = |live: &str, prior: &str, expected: &str, local| {
        if live == expected {
            Ok(Some(prior.to_owned()))
        } else {
            rewrite_table_column(live, table, sheet_name, local, current, caption)
                .map_err(|error| OpError::TableHeaderUnsupported(error.to_string()))
        }
    };
    let mut edits = Vec::new();
    for (owner, at, (prior, expected)) in formulas {
        let live = workbook
            .sheet(*owner)
            .and_then(|sheet| sheet.cell(*at))
            .and_then(|cell| cell.formula.as_deref());
        if let Some(live) = live
            && let Some(formula) = restore(
                live,
                prior,
                expected,
                *owner == *sheet && definition.range.contains(*at),
            )?
        {
            edits.push((*owner, *at, (live.to_owned(), formula)));
        }
    }
    let mut names = Vec::new();
    let mut seen = HashSet::new();
    let mut bytes = defined_names_bytes(&workbook.defined_names);
    for (name, expected) in defined_names {
        let matched = workbook
            .defined_names
            .iter()
            .enumerate()
            .find(|(index, live)| {
                live.local_sheet == name.local_sheet
                    && live.name == name.name
                    && !seen.contains(index)
            });
        if let Some((index, live)) = matched {
            seen.insert(index);
            if let Some(formula) = restore(&live.formula, &name.formula, expected, false)? {
                bytes = bytes - live.formula.len() + formula.len();
                names.push((index, live.clone(), formula));
            }
        }
    }
    if !names.is_empty() && bytes > MAX_DEFINED_NAME_BYTES {
        return Err(OpError::TooManyDefinedNames);
    }
    for (owner, at, (_, formula)) in &edits {
        workbook.sheets[owner.0 as usize].cell_mut(*at).unwrap().formula = Some(formula.clone());
    }
    let defined_names = names
        .into_iter()
        .map(|(index, prior, formula)| {
            workbook.defined_names[index].formula = formula.clone();
            (prior, formula)
        })
        .collect();
    let header = &mut workbook.sheets[sheet.0 as usize];
    let old = header.cell(*at).map(CellState::from).unwrap_or_default();
    header.set_cell(*at, cell.clone().into());
    Ok(Op::RestoreTableColumn {
        sheet: *sheet,
        at: *at,
        table: table.clone(),
        caption: std::mem::replace(&mut workbook.tables[index].columns[column], caption.clone()),
        cell: old,
        formulas: edits,
        defined_names,
    })
}
