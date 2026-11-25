use std::collections::{HashMap, HashSet};

use quick_xml::Reader;
use quick_xml::events::Event;
use xlsx_model::{CellRef, CellValue, Table, Workbook};

use crate::ParseError;
use crate::package::{PreservedPackage, normalized_part_name};
use crate::xml::{xml_err, xstring_attribute_value};

use super::PartStore;

pub(super) fn patch_table_parts(
    workbook: &Workbook,
    package: &PreservedPackage,
    origins: &[Option<usize>],
    parts: &mut PartStore<'_>,
) -> Result<(), ParseError> {
    if package.table_parts.is_empty() || workbook.tables.is_empty() {
        return Ok(());
    }
    let key = |origin, table: &Table| (origin, table.name.to_lowercase());
    let mut tables = HashMap::new();
    let mut ambiguous = HashSet::new();
    for table in &workbook.tables {
        if let Some(origin) = origins.get(table.sheet.0 as usize).copied().flatten() {
            let key = key(origin, table);
            if tables.insert(key.clone(), table).is_some() {
                ambiguous.insert(key);
            }
        }
    }
    let mut source_keys = HashSet::new();
    let mut part_counts = HashMap::new();
    let originals = package
        .original_workbook
        .tables
        .iter()
        .zip(&package.table_parts);
    for (original, path) in originals.clone() {
        let key = key(original.sheet.0 as usize, original);
        if !source_keys.insert(key.clone()) {
            ambiguous.insert(key);
        }
        *part_counts.entry(normalized_part_name(path)).or_insert(0) += 1;
    }
    for (original, path) in originals {
        let key = key(original.sheet.0 as usize, original);
        let Some(table) = tables.get(&key) else {
            continue;
        };
        if original.header_rows == 0 || table.header_rows == 0 {
            continue;
        }
        if ambiguous.contains(&key) {
            return Err(ParseError::UnsupportedEdit(format!(
                "cannot identify the source part for table {} on sheet {}: duplicate display name",
                table.name, table.sheet.0
            )));
        }
        if table.columns.len() != original.columns.len() {
            return Err(column_count_error(path));
        }
        let Some(names) = header_names(workbook, table, original)? else {
            continue;
        };
        if names
            .iter()
            .zip(&original.columns)
            .all(|(name, old)| names_equal(name, old))
        {
            continue;
        }
        if part_counts[&normalized_part_name(path)] > 1 {
            return Err(ParseError::UnsupportedEdit(format!(
                "cannot patch table {}: table part {path} is shared",
                table.name
            )));
        }
        let source = package
            .part_bytes(path)
            .ok_or_else(|| ParseError::MissingPart(path.clone()))?;
        let changed = patch_names(path, source, &original.columns, &names)?;
        parts.set(path.clone(), changed)?;
    }
    Ok(())
}

fn names_equal(left: &str, right: &str) -> bool {
    left.eq_ignore_ascii_case(right)
        || ((!left.is_ascii() || !right.is_ascii()) && left.to_lowercase() == right.to_lowercase())
}

fn header_names<'a>(
    workbook: &'a Workbook,
    table: &Table,
    original: &Table,
) -> Result<Option<Vec<&'a str>>, ParseError> {
    let sheet = &workbook.sheets[table.sheet.0 as usize];
    let start = table.range.start;
    let mut seen = HashSet::new();
    let mut names = Vec::with_capacity(table.columns.len());
    for col in start.col..=table.range.end.col {
        let at = CellRef::new(start.row, col);
        let name = sheet.cell(at).and_then(|cell| match &cell.value {
            CellValue::Text { value } if cell.formula.is_none() && !value.is_empty() => {
                Some(value.as_str())
            }
            _ => None,
        });
        if let Some(name) = name
            && seen.insert(name.to_lowercase())
        {
            names.push(name);
        } else {
            if table.columns == original.columns {
                return Ok(None);
            }
            return Err(ParseError::UnsupportedEdit(format!(
                "table {} column {} ({}) must have a non-empty, unique text header",
                table.name,
                col - start.col + 1,
                at.to_a1()
            )));
        }
    }
    if names.len() != table.columns.len() {
        return Err(column_count_error(&table.name));
    }
    Ok(Some(names))
}

fn column_count_error(path: &str) -> ParseError {
    ParseError::UnsupportedEdit(format!("table column count does not match {path}"))
}

fn patch_names(
    path: &str,
    source: &[u8],
    original: &[String],
    names: &[&str],
) -> Result<Vec<u8>, ParseError> {
    let mut reader = Reader::from_reader(source);
    let mut output = Vec::with_capacity(source.len());
    let mut cursor = 0;
    let mut ordinal = 0;
    let mut depth = 0;
    let mut root = false;
    let mut columns = false;
    loop {
        let element = match reader.read_event().map_err(xml_err)? {
            Event::Start(element) => {
                let is_column =
                    columns && depth == 2 && element.local_name().as_ref() == b"tableColumn";
                if depth == 0 {
                    root = element.local_name().as_ref() == b"table";
                } else if root && depth == 1 {
                    columns = element.local_name().as_ref() == b"tableColumns";
                }
                depth += 1;
                if !is_column {
                    continue;
                }
                element
            }
            Event::Empty(element)
                if columns && depth == 2 && element.local_name().as_ref() == b"tableColumn" =>
            {
                element
            }
            Event::End(_) => {
                depth -= 1;
                if depth == 1 {
                    columns = false;
                }
                continue;
            }
            Event::Eof => break,
            _ => continue,
        };
        let (Some(old), Some(new)) = (original.get(ordinal), names.get(ordinal)) else {
            return Err(column_count_error(path));
        };
        ordinal += 1;
        if names_equal(old, new) {
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
    use std::borrow::Cow;

    use super::*;
    use xlsx_model::{Cell, SheetId};

    const PATH: &str = "xl/tables/table1.xml";
    const TABLE: &str =
        r#"<table displayName="Sales" ref="A1:A2"><tableColumns count="1"><tableColumn name="Old"/></tableColumns></table>"#;

    #[test]
    fn patch_names_keeps_surrounding_bytes() {
        let source = b"\xEF\xBB\xBF<table xmlns:m='urn:test'><tableColumns>\n  \
            <m:tableColumn id='1' name='Old'/>\n  \
            <tableColumn name=\"Other &amp; name\" id=\"2\"><ext/></tableColumn>\n\
            <extLst><tableColumn name='Inside extension'/></extLst>\n\
            </tableColumns><extLst><tableColumns><tableColumn name='Outside columns'/>\
            </tableColumns></extLst></table>";
        let original = vec!["Old".into(), "Other & name".into()];
        let renamed = ["A&'\t_x0041_", "other & Name"];
        let expected = b"\xEF\xBB\xBF<table xmlns:m='urn:test'><tableColumns>\n  \
            <m:tableColumn id='1' name='A&amp;&apos;_x0009__x005F_x0041_'/>\n  \
            <tableColumn name=\"Other &amp; name\" id=\"2\"><ext/></tableColumn>\n\
            <extLst><tableColumn name='Inside extension'/></extLst>\n\
            </tableColumns><extLst><tableColumns><tableColumn name='Outside columns'/>\
            </tableColumns></extLst></table>";
        let parsed = crate::parse_workbook(&crate::tests::package_with_table(
            std::str::from_utf8(source)
                .unwrap()
                .replace("<table ", "<table displayName='Sales' ref='A1:B2' ")
                .as_str(),
        ))
        .unwrap();
        assert_eq!(parsed.tables[0].columns, original);
        let output = patch_names("table.xml", source, &original, &renamed).unwrap();
        assert_eq!(output.as_slice(), expected.as_slice());
        assert!(patch_names("table.xml", source, &original, &renamed[..1]).is_err());
    }

    #[test]
    fn a_shared_part_cannot_be_patched() {
        let mut parsed =
            crate::parse_workbook_with_owned_package(crate::tests::package_with_table(TABLE))
                .unwrap();
        let package = &mut parsed.package;
        let mut shared = package.original_workbook.tables[0].clone();
        shared.sheet = SheetId(1);
        package.original_workbook.tables.push(shared);
        package.table_parts.push(PATH.to_owned());
        parsed.workbook.sheets[0].set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::Text { value: "New".into() },
                ..Cell::default()
            },
        );
        let mut parts = PartStore::new(&package.parts);
        let error =
            patch_table_parts(&parsed.workbook, package, &[Some(0)], &mut parts).unwrap_err();
        assert!(error.to_string().contains("is shared"));
        assert_eq!(parts.get(PATH), Some(TABLE.as_bytes()));
    }

    #[test]
    fn an_invalid_unedited_header_leaves_the_part_byte_identical() {
        let mut parsed =
            crate::parse_workbook_with_owned_package(crate::tests::package_with_table(TABLE))
                .unwrap();
        let package = &parsed.package;
        parsed.workbook.sheets[0].set_cell(
            CellRef::new(1, 0),
            Cell {
                value: CellValue::Number { value: 7.0 },
                ..Cell::default()
            },
        );
        let mut parts = PartStore::new(&package.parts);
        patch_table_parts(&parsed.workbook, package, &[Some(0)], &mut parts).unwrap();
        assert!(parts.finish().iter().any(|(path, bytes)| {
            path == PATH && matches!(bytes, Cow::Borrowed(bytes) if *bytes == TABLE.as_bytes())
        }));
    }
}
