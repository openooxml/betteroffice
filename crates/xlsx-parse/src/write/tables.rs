use std::collections::{HashMap, HashSet};

use quick_xml::Reader;
use quick_xml::events::Event;
use xlsx_model::{CellRef, CellValue, Table, Workbook};

use crate::ParseError;
use crate::package::PreservedPackage;
use crate::xml::{xml_err, xstring_attribute_value};

use super::PartStore;

pub(super) fn patch_table_parts(
    workbook: &Workbook,
    package: &PreservedPackage,
    origins: &[Option<usize>],
    parts: &mut PartStore<'_>,
) -> Result<(), ParseError> {
    let tables = workbook
        .tables
        .iter()
        .filter_map(|table| {
            let origin = origins.get(table.sheet.0 as usize).copied().flatten()?;
            Some(((origin, table.name.as_str()), table))
        })
        .collect::<HashMap<_, _>>();
    for (original, path) in package.original_workbook.tables.iter().zip(&package.table_parts) {
        let Some(table) = tables.get(&(original.sheet.0 as usize, original.name.as_str())) else {
            continue;
        };
        if table.columns.len() != original.columns.len() {
            return Err(column_count_error(path));
        }
        validate_headers(workbook, table, &package.original_workbook, original, path)?;
        if table.columns == original.columns {
            continue;
        }
        let mut names = HashSet::new();
        if table
            .columns
            .iter()
            .any(|name| name.is_empty() || !names.insert(name.to_lowercase()))
        {
            return Err(ParseError::UnsupportedEdit(format!(
                "table column names in {path} must be non-empty and unique"
            )));
        }
        let source = package
            .part_bytes(path)
            .ok_or_else(|| ParseError::MissingPart(path.clone()))?;
        let changed = patch_names(path, source, &original.columns, &table.columns)?;
        parts.set(path.clone(), changed)?;
    }
    Ok(())
}

fn validate_headers(
    workbook: &Workbook,
    table: &Table,
    source: &Workbook,
    original: &Table,
    path: &str,
) -> Result<(), ParseError> {
    if table.header_rows == 0 {
        return Ok(());
    }
    let sheet = &workbook.sheets[table.sheet.0 as usize];
    let source_sheet = &source.sheets[original.sheet.0 as usize];
    let start = table.range.start;
    let source_start = original.range.start;
    for (col, (name, old_name)) in
        (start.col..=table.range.end.col).zip(table.columns.iter().zip(&original.columns))
    {
        let cell = sheet.cell(CellRef::new(start.row, col));
        let source_col = source_start.col + (col - start.col);
        let old = source_sheet.cell(CellRef::new(source_start.row, source_col));
        let formula = cell.and_then(|cell| cell.formula.as_deref());
        let value = cell.map_or(&CellValue::Empty, |cell| &cell.value);
        if name == old_name
            && formula == old.and_then(|cell| cell.formula.as_deref())
            && (formula.is_some() || value == old.map_or(&CellValue::Empty, |cell| &cell.value))
        {
            continue;
        }
        if formula.is_some()
            || !matches!(value, CellValue::Text { value } if value == name && !value.is_empty())
        {
            return Err(ParseError::UnsupportedEdit(format!(
                "table header cell does not match its column name in {path}"
            )));
        }
    }
    Ok(())
}

fn column_count_error(path: &str) -> ParseError {
    ParseError::UnsupportedEdit(format!("table column count does not match {path}"))
}

fn patch_names(
    path: &str,
    source: &[u8],
    original: &[String],
    names: &[String],
) -> Result<Vec<u8>, ParseError> {
    let mut reader = Reader::from_reader(source);
    let mut output = Vec::with_capacity(source.len());
    let mut cursor = 0;
    let mut ordinal = 0;
    loop {
        let element = match reader.read_event().map_err(xml_err)? {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"tableColumn" =>
            {
                element
            }
            Event::Eof => break,
            _ => continue,
        };
        let (Some(old), Some(new)) = (original.get(ordinal), names.get(ordinal)) else {
            return Err(column_count_error(path));
        };
        ordinal += 1;
        if old == new {
            continue;
        }
        let name = element
            .try_get_attribute(b"name")
            .map_err(xml_err)?
            .ok_or_else(|| ParseError::Malformed(format!("missing table column name in {path}")))?;
        let start = name
            .value
            .as_ptr()
            .addr()
            .checked_sub(source.as_ptr().addr())
            .filter(|start| {
                *start >= cursor
                    && source
                        .get(*start..)
                        .is_some_and(|tail| tail.starts_with(&name.value))
            })
            .ok_or_else(|| {
                ParseError::Malformed(format!("invalid table column name span in {path}"))
            })?;
        output.extend_from_slice(&source[cursor..start]);
        output.extend_from_slice(xstring_attribute_value(new).as_bytes());
        cursor = start + name.value.len();
    }
    if ordinal != names.len() {
        return Err(column_count_error(path));
    }
    output.extend_from_slice(&source[cursor..]);
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn patch_names_keeps_surrounding_bytes() {
        let source = b"\xEF\xBB\xBF<table xmlns:m='urn:test'><tableColumns>\n  \
            <m:tableColumn id='1' name='Old'/>\n  \
            <tableColumn name=\"Other &amp; name\" id=\"2\"><ext/></tableColumn>\n\
            </tableColumns></table>";
        let original = vec!["Old".into(), "Other & name".into()];
        let renamed = vec!["A&'\t_x0041_".into(), "Other & name".into()];
        let expected = b"\xEF\xBB\xBF<table xmlns:m='urn:test'><tableColumns>\n  \
            <m:tableColumn id='1' name='A&amp;&apos;_x0009__x005F_x0041_'/>\n  \
            <tableColumn name=\"Other &amp; name\" id=\"2\"><ext/></tableColumn>\n\
            </tableColumns></table>";
        let output = patch_names("table.xml", source, &original, &renamed).unwrap();
        assert_eq!(output.as_slice(), expected.as_slice());
        assert!(patch_names("table.xml", source, &original, &renamed[..1]).is_err());
    }
}
