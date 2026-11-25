use std::collections::{HashMap, HashSet};

use quick_xml::Reader;
use quick_xml::events::Event;
use xlsx_model::{CellRef, CellValue, Table, Workbook};

use crate::ParseError;
use crate::package::{PreservedPackage, normalized_part_name};
use crate::xml::{xml_err, xstring_attribute_value};

use super::PartStore;

pub(super) fn changed_tables<'a>(
    workbook: &'a Workbook,
    package: &'a PreservedPackage,
    origins: &[Option<usize>],
) -> Vec<(&'a Table, &'a Table, &'a str)> {
    let originals = &package.original_workbook.tables;
    if originals.is_empty()
        || workbook.tables.is_empty()
        || (workbook.tables.len() == originals.len()
            && workbook
                .tables
                .iter()
                .zip(originals)
                .all(|(table, source)| {
                    origins.get(table.sheet.0 as usize).copied().flatten()
                        == Some(source.sheet.0 as usize)
                        && table.name == source.name
                        && table.columns == source.columns
                }))
    {
        return Vec::new();
    }
    let mut sources = HashMap::<_, Vec<_>>::new();
    for (source, path) in originals.iter().zip(&package.table_parts).rev() {
        sources
            .entry((source.sheet.0 as usize, source.name.to_lowercase()))
            .or_default()
            .push((source, path.as_str()));
    }
    workbook
        .tables
        .iter()
        .filter_map(|table| {
            let origin = origins.get(table.sheet.0 as usize).copied().flatten()?;
            let matches = sources.get_mut(&(origin, table.name.to_lowercase()))?;
            let same_range = matches
                .iter()
                .rposition(|(source, _)| source.range == table.range);
            let (source, path) = match same_range {
                Some(index) => matches.remove(index),
                None => matches.pop()?,
            };
            (table.columns != source.columns).then_some((table, source, path))
        })
        .collect()
}

pub(super) fn patch_table_parts(
    workbook: &Workbook,
    package: &PreservedPackage,
    changes: &[(&Table, &Table, &str)],
    parts: &mut PartStore<'_>,
) -> Result<(), ParseError> {
    if changes.is_empty() {
        return Ok(());
    }
    let mut destinations = HashMap::new();
    for (path, bytes) in &package.parts {
        destinations
            .entry(normalized_part_name(path))
            .and_modify(|entry| *entry = None)
            .or_insert(Some((path, bytes)));
    }
    let originals = &package.original_workbook.tables;
    for &(table, source, path) in changes {
        let name = table.name.to_lowercase();
        let duplicate = |tables: &[Table]| {
            tables
                .iter()
                .filter(|table| table.name.to_lowercase() == name)
                .nth(1)
                .is_some()
        };
        if duplicate(originals) || duplicate(&workbook.tables) {
            return Err(ParseError::UnsupportedEdit(format!(
                "cannot patch table {} in {path}: duplicate display name",
                table.name
            )));
        }
        if table.columns.len() != source.columns.len() {
            return Err(column_count_error(path));
        }
        let normalized = normalized_part_name(path);
        if package
            .table_parts
            .iter()
            .filter(|path| {
                path.trim_start_matches('/')
                    .eq_ignore_ascii_case(&normalized)
            })
            .nth(1)
            .is_some()
        {
            return Err(ParseError::UnsupportedEdit(format!(
                "cannot patch table {}: table part {path} is shared",
                table.name
            )));
        }
        let destination = destinations
            .get(&normalized)
            .ok_or_else(|| ParseError::MissingPart(path.to_owned()))?;
        let Some((destination, bytes)) = destination else {
            return Err(ParseError::UnsupportedEdit(format!(
                "cannot patch table {}: table part {path} has multiple destinations",
                table.name
            )));
        };
        validate_headers(workbook, table)?;
        parts.set(
            (*destination).clone(),
            patch_names(path, bytes, &source.columns, &table.columns)?,
        )?;
    }
    Ok(())
}

fn validate_headers(workbook: &Workbook, table: &Table) -> Result<(), ParseError> {
    let sheet = &workbook.sheets[table.sheet.0 as usize];
    let start = table.range.start;
    let width = table.range.end.col - start.col + 1;
    if table.header_rows == 0 || table.columns.len() != width as usize {
        return Err(ParseError::UnsupportedEdit(format!(
            "table {} does not have a header cell for each column",
            table.name
        )));
    }
    let mut seen = HashSet::new();
    for (col, caption) in (start.col..=table.range.end.col).zip(&table.columns) {
        let at = CellRef::new(start.row, col);
        if caption.is_empty()
            || !seen.insert(caption.to_lowercase())
            || !sheet.cell(at).is_some_and(|cell| {
                cell.formula.is_none()
                    && matches!(&cell.value, CellValue::Text { value } if value == caption)
            })
        {
            return Err(ParseError::UnsupportedEdit(format!(
                "table {} column {} ({}) needs non-empty, unique text matching its caption",
                table.name,
                col - start.col + 1,
                at.to_a1()
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
    use xlsx_model::{Cell, SheetId};

    const PATH: &str = "xl/tables/table1.xml";
    const TABLE: &str = r#"<table displayName="Sales" ref="A1:A2"><tableColumns count="1"><tableColumn name="Old"/></tableColumns></table>"#;

    #[test]
    fn patch_names_keeps_surrounding_bytes() {
        let source = b"\xEF\xBB\xBF<table xmlns:m='urn:test'><tableColumns>\n  \
            <m:tableColumn id='1' name='Old'/>\n  \
            <tableColumn name=\"Other &amp; name\" id=\"2\"><ext/></tableColumn>\n\
            <extLst><tableColumn name='Inside extension'/></extLst>\n\
            </tableColumns><extLst><tableColumns><tableColumn name='Outside columns'/>\
            </tableColumns></extLst></table>";
        let original = vec!["Old".into(), "Other & name".into()];
        let renamed = ["A&'\t_x0041_".into(), "Other & name".into()];
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
        shared.name = "Other".into();
        package.original_workbook.tables.push(shared);
        package.table_parts.push(PATH.to_owned());
        parsed.workbook.tables[0].columns[0] = "New".into();
        let mut parts = PartStore::new(&package.parts);
        let changes = changed_tables(&parsed.workbook, package, &[Some(0)]);
        let error = patch_table_parts(&parsed.workbook, package, &changes, &mut parts).unwrap_err();
        assert!(error.to_string().contains("is shared"));
        assert_eq!(parts.get(PATH), Some(TABLE.as_bytes()));
    }

    #[test]
    fn case_variant_parts_are_ambiguous_only_if_patched() {
        let mut parsed =
            crate::parse_workbook_with_owned_package(crate::tests::package_with_table(TABLE))
                .unwrap();
        let package = &mut parsed.package;
        package
            .parts
            .push((PATH.to_ascii_uppercase(), TABLE.as_bytes().to_vec()));
        assert!(changed_tables(&parsed.workbook, package, &[Some(0)]).is_empty());
        parsed.workbook.tables[0].columns[0] = "New".into();
        let mut parts = PartStore::new(&package.parts);
        let changes = changed_tables(&parsed.workbook, package, &[Some(0)]);
        let error = patch_table_parts(&parsed.workbook, package, &changes, &mut parts).unwrap_err();
        assert!(error.to_string().contains("multiple destinations"));
    }

    #[test]
    fn a_changed_caption_requires_an_exact_match_in_the_header_cell() {
        let mut parsed =
            crate::parse_workbook_with_owned_package(crate::tests::package_with_table(TABLE))
                .unwrap();
        parsed.workbook.tables[0].columns[0] = "New".into();
        parsed.workbook.sheets[0].set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::Text {
                    value: "new".into(),
                },
                ..Cell::default()
            },
        );
        let package = &parsed.package;
        let changes = changed_tables(&parsed.workbook, package, &[Some(0)]);
        let error = patch_table_parts(
            &parsed.workbook,
            package,
            &changes,
            &mut PartStore::new(&package.parts),
        )
        .unwrap_err();
        assert!(error.to_string().contains("matching its caption"));
    }
}
