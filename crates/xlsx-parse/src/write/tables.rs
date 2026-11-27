use std::collections::{HashMap, HashSet};

use quick_xml::Reader;
use quick_xml::events::Event;
use xlsx_calc::lexer::{may_reference_table_column, rewrite_table_columns};
use xlsx_model::{CellRef, CellValue, Table, Workbook};

use crate::ParseError;
use crate::package::{PreservedPackage, normalized_part_name};
use crate::xml::{
    collect_text, formula_text, next_event, reader, utf8_bom_len, xml_err, xstring_attribute_value,
};

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
    table_pairs(workbook, package, origins)
        .into_iter()
        .filter(|(table, source, _)| table.columns != source.columns)
        .collect()
}

fn table_pairs<'a>(
    workbook: &'a Workbook,
    package: &'a PreservedPackage,
    origins: &[Option<usize>],
) -> Vec<(&'a Table, &'a Table, &'a str)> {
    let mut sources = HashMap::<_, Vec<_>>::new();
    for (source, path) in package
        .original_workbook
        .tables
        .iter()
        .zip(&package.table_parts)
        .rev()
    {
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
            Some((table, source, path))
        })
        .collect()
}

impl PreservedPackage {
    #[doc(hidden)]
    pub fn validate_table_column_rename(
        &self,
        workbook: &Workbook,
        origins: &[Option<usize>],
        table: &Table,
        column: usize,
        caption: &str,
    ) -> Result<(), ParseError> {
        let pairs = table_pairs(workbook, self, origins);
        let (_, source, path) = pairs
            .iter()
            .copied()
            .find(|(target, _, _)| std::ptr::eq(*target, table))
            .ok_or_else(|| {
                ParseError::UnsupportedEdit(format!(
                    "cannot find source part for table {}",
                    table.name
                ))
            })?;
        let mut columns = table.columns.clone();
        *columns
            .get_mut(column)
            .ok_or_else(|| column_count_error(path))? = caption.to_owned();
        if columns != source.columns {
            let (path, bytes) =
                self.validate_table_change(workbook, table, source, path, &columns, Some(column))?;
            patch_names(path, bytes, &source.columns, &columns)?;
        }
        let renames = caption_renames(self, &pairs, Some((table, column, caption)));
        self.patch_table_formulas(&renames, |_, _| Ok(()))
    }

    fn validate_table_change(
        &self,
        workbook: &Workbook,
        table: &Table,
        source: &Table,
        path: &str,
        columns: &[String],
        proposed_column: Option<usize>,
    ) -> Result<(&str, &[u8]), ParseError> {
        let destination = self.table_part_destination(path, &table.name)?;
        let name = table.name.to_lowercase();
        let duplicate = |tables: &[Table]| {
            tables
                .iter()
                .filter(|table| table.name.to_lowercase() == name)
                .nth(1)
                .is_some()
        };
        if duplicate(&self.original_workbook.tables) || duplicate(&workbook.tables) {
            return Err(ParseError::UnsupportedEdit(format!(
                "cannot patch table {} in {path}: duplicate display name",
                table.name
            )));
        }
        if columns.len() != source.columns.len() {
            return Err(column_count_error(path));
        }
        validate_headers(workbook, table, columns, &source.columns, proposed_column)?;
        Ok(destination)
    }

    fn table_part_destination(&self, path: &str, table: &str) -> Result<(&str, &[u8]), ParseError> {
        reject_shared_part(self, path, table)?;
        let normalized = normalized_part_name(path);
        let mut destinations = self.parts.iter().filter(|(path, _)| {
            path.trim_start_matches('/')
                .eq_ignore_ascii_case(&normalized)
        });
        let (destination, bytes) = destinations
            .next()
            .ok_or_else(|| ParseError::MissingPart(path.to_owned()))?;
        if destinations.next().is_some() {
            return Err(ParseError::UnsupportedEdit(format!(
                "cannot patch table {table}: table part {path} has multiple destinations"
            )));
        }
        Ok((destination, bytes))
    }

    fn patch_table_formulas(
        &self,
        renames: &[CaptionRename<'_>],
        mut patched: impl FnMut(&str, Vec<u8>) -> Result<(), ParseError>,
    ) -> Result<(), ParseError> {
        if renames.is_empty() {
            return Ok(());
        }
        for (table, path) in self.original_workbook.tables.iter().zip(&self.table_parts) {
            let bytes = self
                .part_bytes(path)
                .ok_or_else(|| ParseError::MissingPart(path.clone()))?;
            if let Some(output) = patch_formulas(path, bytes, table, renames)? {
                let (destination, _) = self.table_part_destination(path, &table.name)?;
                patched(destination, output)?;
            }
        }
        Ok(())
    }
}

struct CaptionRename<'a> {
    table: &'a Table,
    sheet: &'a str,
    old: &'a str,
    new: &'a str,
}

fn caption_renames<'a>(
    package: &'a PreservedPackage,
    pairs: &[(&'a Table, &'a Table, &str)],
    proposal: Option<(&Table, usize, &'a str)>,
) -> Vec<CaptionRename<'a>> {
    let mut renames = Vec::new();
    for &(table, source, _) in pairs {
        let sheet = &package.original_workbook.sheets[source.sheet.0 as usize].name;
        for (index, (old, current)) in source.columns.iter().zip(&table.columns).enumerate() {
            let new = proposal
                .filter(|(target, column, _)| std::ptr::eq(table, *target) && *column == index)
                .map_or(current.as_str(), |(_, _, caption)| caption);
            if old.to_lowercase() != new.to_lowercase() {
                renames.push(CaptionRename {
                    table: source,
                    sheet,
                    old,
                    new,
                });
            }
        }
    }
    renames
}

fn patch_formulas(
    path: &str,
    bytes: &[u8],
    table: &Table,
    renames: &[CaptionRename<'_>],
) -> Result<Option<Vec<u8>>, ParseError> {
    let mut xml = reader(bytes);
    let bom = utf8_bom_len(bytes);
    let mut buffer = Vec::new();
    let mut depth = 0;
    let mut root = false;
    let mut columns = false;
    let mut column = false;
    let mut output = Vec::new();
    let mut cursor = 0;
    loop {
        match next_event(&mut xml, &mut buffer, &mut depth)? {
            Event::Start(element) => {
                let name = element.local_name();
                let name = name.as_ref();
                match depth {
                    1 => root = name == b"table",
                    2 => columns = root && name == b"tableColumns",
                    3 => column = columns && name == b"tableColumn",
                    _ => {}
                }
                if !matches!(name, b"calculatedColumnFormula" | b"totalsRowFormula") {
                    continue;
                }
                let direct = depth == 4 && column;
                let start = bom + xml.buffer_position() as usize;
                let formula = collect_text(&mut xml, &mut buffer, &mut depth)?;
                if !renames.iter().any(|rename| {
                    may_reference_table_column(
                        &formula,
                        &rename.table.name,
                        std::ptr::eq(table, rename.table),
                        rename.old,
                    )
                }) {
                    continue;
                }
                let rewritten = rewrite_table_columns(&formula, |scope, name, caption| {
                    renames.iter().find_map(|rename| {
                        let matches = if name.is_empty() {
                            std::ptr::eq(table, rename.table)
                        } else {
                            name.to_lowercase() == rename.table.name.to_lowercase()
                                && scope.is_none_or(|scope| {
                                    scope.to_lowercase() == rename.sheet.to_lowercase()
                                })
                        };
                        (matches && caption.to_lowercase() == rename.old.to_lowercase())
                            .then_some(rename.new)
                    })
                })
                .map_err(|error| {
                    ParseError::UnsupportedEdit(format!(
                        "cannot rewrite formula in {path}: {error}"
                    ))
                })?;
                if let Some(rewritten) = rewritten {
                    let tail = &bytes[start..bom + xml.buffer_position() as usize];
                    let end = tail.iter().rposition(|&byte| byte == b'<');
                    let Some(end) = end.filter(|&end| direct && !tail[..end].contains(&b'<'))
                    else {
                        return Err(ParseError::UnsupportedEdit(format!(
                            "cannot patch formula text in {path}"
                        )));
                    };
                    output.extend_from_slice(&bytes[cursor..start]);
                    output.extend_from_slice(formula_text(&rewritten).as_bytes());
                    cursor = start + end;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok((cursor != 0).then(|| {
        output.extend_from_slice(&bytes[cursor..]);
        output
    }))
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
    let mut names = HashMap::new();
    for &(table, source, path) in changes {
        let (destination, bytes) =
            package.validate_table_change(workbook, table, source, path, &table.columns, None)?;
        names.insert(
            normalized_part_name(path),
            (destination, bytes, source, table),
        );
    }
    let renames = caption_renames(package, changes, None);
    package.patch_table_formulas(&renames, |path, mut bytes| {
        if let Some((_, _, source, table)) = names.remove(&normalized_part_name(path)) {
            bytes = patch_names(path, &bytes, &source.columns, &table.columns)?;
        }
        parts.set(path.to_owned(), bytes)
    })?;
    for (_, (path, bytes, source, table)) in names {
        parts.set(
            path.to_owned(),
            patch_names(path, bytes, &source.columns, &table.columns)?,
        )?;
    }
    Ok(())
}

fn reject_shared_part(
    package: &PreservedPackage,
    path: &str,
    table_name: &str,
) -> Result<(), ParseError> {
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
            "cannot patch table {table_name}: table part {path} is shared"
        )));
    }
    Ok(())
}

fn validate_headers(
    workbook: &Workbook,
    table: &Table,
    columns: &[String],
    original: &[String],
    proposed_column: Option<usize>,
) -> Result<(), ParseError> {
    let sheet = &workbook.sheets[table.sheet.0 as usize];
    let start = table.range.start;
    let width = table.range.end.col - start.col + 1;
    if table.header_rows == 0 || columns.len() != width as usize {
        return Err(ParseError::UnsupportedEdit(format!(
            "table {} does not have a header cell for each column",
            table.name
        )));
    }
    let mut seen = HashSet::new();
    for (index, (caption, original)) in columns.iter().zip(original).enumerate() {
        let at = CellRef::new(start.row, start.col + index as u32);
        let folded = caption.to_lowercase();
        let matches = proposed_column == Some(index)
            || sheet.cell(at).is_some_and(|cell| {
                cell.formula.is_none()
                    && matches!(
                        &cell.value,
                        CellValue::Text { value }
                            if value == caption
                                || caption == original && value.to_lowercase() == folded
                    )
            });
        if caption.is_empty() || !matches || !seen.insert(folded) {
            return Err(ParseError::UnsupportedEdit(format!(
                "table {} column {} ({}) needs non-empty, unique text matching its caption",
                table.name,
                index + 1,
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
