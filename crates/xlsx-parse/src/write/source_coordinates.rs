use std::ops::Range;

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, Writer};
use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::{CellRange, CellRef};

use crate::axis::{AxisMap, SheetAxes};
use crate::formula::{Reference, ReferenceStatus, classify_references, sheet_name};
use crate::package::attributes;
use crate::xml::{attr, resolve_entity, xml_err};
use crate::{MAX_DEPTH, ParseError};

pub(super) fn unchanged_with_defined_names(
    source: &[u8],
    axes: Option<&SheetAxes>,
    sheet_axes: &[(&str, Option<&SheetAxes>)],
    axes_changed: bool,
    defined_names: Option<&[&str]>,
) -> bool {
    (axes.is_none_or(SheetAxes::is_identity) && !axes_changed)
        || inspect(source, axes, sheet_axes, defined_names).unwrap_or(false)
}

#[cfg(test)]
fn unchanged(
    source: &[u8],
    axes: Option<&SheetAxes>,
    sheet_axes: &[(&str, Option<&SheetAxes>)],
    axes_changed: bool,
) -> bool {
    unchanged_with_defined_names(source, axes, sheet_axes, axes_changed, Some(&[]))
}

fn fixed(axis: &AxisMap, range: Range<u32>) -> bool {
    range.is_empty() || axis.current_ranges_bounded(range.clone()) == [range]
}

fn references(value: &str, axes: &SheetAxes) -> bool {
    let mut ranges = value.split_whitespace().peekable();
    ranges.peek().is_some()
        && ranges.all(|value| {
            CellRange::parse_a1(value).is_ok_and(|range| {
                fixed(&axes.rows, range.start.row..range.end.row + 1)
                    && fixed(&axes.cols, range.start.col..range.end.col + 1)
            })
        })
}

fn row_index(element: &BytesStart<'_>, current: Option<u32>) -> Result<Option<u32>, ParseError> {
    Ok(match attr(element, b"r")? {
        Some(value) => value
            .parse::<u32>()
            .ok()
            .filter(|index| (1..=MAX_ROWS).contains(index))
            .map(|index| index - 1),
        None => Some(current.map_or(0, |row| row + 1)),
    })
}

fn formula(
    value: &str,
    axes: Option<&SheetAxes>,
    sheet_axes: &[(&str, Option<&SheetAxes>)],
    defined_names: Option<&[&str]>,
) -> ReferenceStatus {
    classify_references(value, defined_names, |sheet, reference| {
        let axes = match sheet {
            None => axes,
            Some(name) => {
                let name = sheet_name(name);
                let mut candidates = sheet_axes
                    .iter()
                    .filter(|(sheet, _)| sheet.eq_ignore_ascii_case(&name));
                let axes = candidates.next().and_then(|(_, axes)| *axes);
                if candidates.next().is_some() {
                    return ReferenceStatus::Unresolvable;
                }
                axes
            }
        };
        let Some(axes) = axes else {
            return ReferenceStatus::Unresolvable;
        };
        let fixed = match reference {
            Reference::Cells(range) => {
                fixed(&axes.rows, range.start.row..range.end.row + 1)
                    && fixed(&axes.cols, range.start.col..range.end.col + 1)
            }
            Reference::Rows(range) => fixed(&axes.rows, range),
            Reference::Columns(range) => fixed(&axes.cols, range),
        };
        if fixed {
            ReferenceStatus::Unmoved
        } else {
            ReferenceStatus::Moved
        }
    })
}

fn text_coordinates(name: &[u8], value: &str, axes: &SheetAxes) -> bool {
    match name {
        b"row" => value
            .trim()
            .parse::<u32>()
            .is_ok_and(|index| axes.rows.current(index) == Some(index)),
        b"col" => value
            .trim()
            .parse::<u32>()
            .is_ok_and(|index| axes.cols.current(index) == Some(index)),
        _ => references(value, axes),
    }
}

fn inspect(
    source: &[u8],
    axes: Option<&SheetAxes>,
    sheet_axes: &[(&str, Option<&SheetAxes>)],
    defined_names: Option<&[&str]>,
) -> Result<bool, ParseError> {
    let formula_axes = axes;
    let identity = SheetAxes::default();
    let axes = axes.unwrap_or(&identity);
    let mut reader = Reader::from_reader(source);
    reader.config_mut().expand_empty_elements = true;
    let mut parents: Vec<Vec<u8>> = Vec::new();
    let mut row = None;
    let mut col = 0;
    let mut positioned_text: Option<(Vec<u8>, String)> = None;
    loop {
        let event = reader.read_event().map_err(xml_err)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                if positioned_text.is_some() {
                    return Ok(false);
                }
                let name = element.local_name();
                let name = name.as_ref();
                let parent = parents.last().map(Vec::as_slice);
                let source_row = name == b"row" && parent == Some(b"sheetData".as_slice());
                let source_cell = name == b"c" && parent == Some(b"row".as_slice());
                if source_row {
                    let Some(index) = row_index(element, row)? else {
                        return Ok(false);
                    };
                    row = Some(index);
                    col = 0;
                    if axes.rows.current(index) != Some(index) {
                        return Ok(false);
                    }
                }
                if source_cell {
                    let at = match attr(element, b"r")? {
                        Some(value) => CellRef::parse_a1(&value).ok(),
                        None => Some(CellRef::new(row.unwrap_or(0), col)),
                    };
                    let Some(at) = at else {
                        return Ok(false);
                    };
                    col = at.col;
                    if axes.rows.current(at.row) != Some(at.row)
                        || axes.cols.current(at.col) != Some(at.col)
                    {
                        return Ok(false);
                    }
                }
                if name == b"col" && parent == Some(b"cols".as_slice()) {
                    let min = attr(element, b"min")?.and_then(|v| v.parse::<u32>().ok());
                    let max = attr(element, b"max")?.and_then(|v| v.parse::<u32>().ok());
                    let (Some(min), Some(max)) = (min, max) else {
                        return Ok(false);
                    };
                    if min == 0 || min > max || max > MAX_COLS || !fixed(&axes.cols, min - 1..max) {
                        return Ok(false);
                    }
                }
                if name == b"brk" && matches!(parent, Some(b"rowBreaks" | b"colBreaks")) {
                    let axis = match parent {
                        Some(b"rowBreaks") => &axes.cols,
                        Some(b"colBreaks") => &axes.rows,
                        _ => return Ok(false),
                    };
                    let min = attr(element, b"min")?.map_or(Some(0), |v| v.parse::<u32>().ok());
                    let limit = if parent == Some(b"rowBreaks".as_slice()) {
                        MAX_COLS
                    } else {
                        MAX_ROWS
                    };
                    let max =
                        attr(element, b"max")?.map_or(Some(limit - 1), |v| v.parse::<u32>().ok());
                    let (Some(min), Some(max)) = (min, max) else {
                        return Ok(false);
                    };
                    let Some(end) = max.checked_add(1).filter(|_| min <= max) else {
                        return Ok(false);
                    };
                    if !fixed(axis, min..end) {
                        return Ok(false);
                    }
                }
                for attribute in attributes(element)? {
                    if attribute.name == "xmlns" || attribute.name.starts_with("xmlns:") {
                        continue;
                    }
                    let key = attribute.name.rsplit(':').next().unwrap_or_default();
                    let value = attribute.value.as_str();
                    let stable = match (name, key) {
                        (b"row" | b"c", "r") if source_row || source_cell => true,
                        (b"row", "r") => row_index(element, None)?
                            .is_some_and(|index| axes.rows.current(index) == Some(index)),
                        (b"c", "r") => references(value, axes),
                        (
                            b"dimension" | b"mergeCell" | b"hyperlink" | b"autoFilter"
                            | b"sortState" | b"sortCondition" | b"f",
                            "ref",
                        )
                        | (b"f", "r1" | "r2")
                        | (b"selection", "activeCell" | "sqref")
                        | (b"pane" | b"sheetView", "topLeftCell")
                        | (
                            b"dataValidation"
                            | b"conditionalFormatting"
                            | b"protectedRange"
                            | b"ignoredError",
                            "sqref",
                        )
                        | (b"cellWatch" | b"inputCells", "r") => references(value, axes),
                        (b"row", "spans") if source_row => value.split_whitespace().all(|span| {
                            let Some((min, max)) = span.split_once(':') else {
                                return false;
                            };
                            let (Ok(min), Ok(max)) = (min.parse::<u32>(), max.parse::<u32>())
                            else {
                                return false;
                            };
                            min > 0
                                && min <= max
                                && max <= MAX_COLS
                                && fixed(&axes.cols, min - 1..max)
                        }),
                        (b"pane", "xSplit" | "ySplit") => {
                            let frozen = attr(element, b"state")?
                                .is_some_and(|state| state == "frozen" || state == "frozenSplit");
                            !frozen
                                || value.parse::<f64>().is_ok_and(|end| {
                                    let (axis, limit) = if key == "xSplit" {
                                        (&axes.cols, MAX_COLS)
                                    } else {
                                        (&axes.rows, MAX_ROWS)
                                    };
                                    end.is_finite()
                                        && end >= 0.0
                                        && end.fract() == 0.0
                                        && end <= f64::from(limit)
                                        && (end == 0.0
                                            || fixed(axis, 0..(end as u32 + 1).min(limit)))
                                })
                        }
                        (b"brk", "id" | "min" | "max")
                            if matches!(parent, Some(b"rowBreaks" | b"colBreaks")) =>
                        {
                            let axis = match (parent, key) {
                                (Some(b"rowBreaks"), "id")
                                | (Some(b"colBreaks"), "min" | "max") => &axes.rows,
                                (Some(b"colBreaks"), "id")
                                | (Some(b"rowBreaks"), "min" | "max") => &axes.cols,
                                _ => return Ok(false),
                            };
                            value
                                .parse::<u32>()
                                .is_ok_and(|index| axis.current(index) == Some(index))
                        }
                        _ => true,
                    };
                    if !stable {
                        return Ok(false);
                    }
                }
                if matches!(event, Event::Start(_)) {
                    if parents.iter().any(|parent| parent == b"extLst")
                        && (matches!(name, b"sqref" | b"ref" | b"f")
                            || (matches!(name, b"row" | b"col")
                                && matches!(parent, Some(b"anchor" | b"from" | b"to"))))
                    {
                        positioned_text = Some((name.to_vec(), String::new()));
                    }
                    parents.push(name.to_vec());
                    if parents.len() > MAX_DEPTH {
                        return Err(ParseError::DepthExceeded);
                    }
                }
            }
            Event::Text(text) => {
                if let Some((_, value)) = &mut positioned_text {
                    value.push_str(&text.decode().map_err(xml_err)?);
                }
            }
            Event::CData(text) => {
                if let Some((_, value)) = &mut positioned_text {
                    value.push_str(&text.decode().map_err(xml_err)?);
                }
            }
            Event::GeneralRef(reference) => {
                if let Some((_, value)) = &mut positioned_text {
                    value.push_str(&resolve_entity(&reference.decode().map_err(xml_err)?)?);
                }
            }
            Event::End(element) => {
                if let Some((name, value)) = positioned_text.take() {
                    let stable = if name == b"f" {
                        formula(&value, formula_axes, sheet_axes, defined_names)
                            == ReferenceStatus::Unmoved
                    } else {
                        text_coordinates(&name, &value, axes)
                    };
                    if !stable {
                        return Ok(false);
                    }
                }
                match (
                    element.local_name().as_ref(),
                    parents.iter().rev().nth(1).map(Vec::as_slice),
                ) {
                    (b"row", Some(b"sheetData")) => row = None,
                    (b"c", Some(b"row")) => col += 1,
                    _ => {}
                }
                parents.pop();
            }
            Event::Eof => return Ok(true),
            _ => {}
        }
    }
}

pub(super) fn explicit_rows(
    source: &[u8],
    axes: &SheetAxes,
) -> Result<Option<Vec<u8>>, ParseError> {
    if axes.is_identity() {
        return Ok(None);
    }
    let mut reader = Reader::from_reader(source);
    let mut out = Writer::new(Vec::new());
    let mut depth = 0;
    let mut cursor = 0;
    let mut changed = false;
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(xml_err)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element)
                if depth == 1 && element.local_name().as_ref() == b"row" =>
            {
                if attr(element, b"r")?.is_none() {
                    let Some(index) = row_index(element, None)? else {
                        return Ok(None);
                    };
                    let index = (index + 1).to_string();
                    out.get_mut().extend_from_slice(&source[cursor..before]);
                    let mut element = element.clone();
                    element.push_attribute(("r", index.as_str()));
                    out.write_event(if matches!(event, Event::Empty(_)) {
                        Event::Empty(element)
                    } else {
                        Event::Start(element)
                    })
                    .map_err(xml_err)?;
                    cursor = reader.buffer_position() as usize;
                    changed = true;
                }
            }
            Event::Eof => {
                if !changed {
                    return Ok(None);
                }
                out.get_mut().extend_from_slice(&source[cursor..]);
                return Ok(Some(out.into_inner()));
            }
            _ => {}
        }
        match event {
            Event::Start(_) => {
                depth += 1;
                if depth > MAX_DEPTH {
                    return Err(ParseError::DepthExceeded);
                }
            }
            Event::End(_) => depth = depth.saturating_sub(1),
            _ => {}
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn unchanged(source: &[u8], axes: &SheetAxes) -> bool {
        super::unchanged(source, Some(axes), &[], !axes.is_identity())
    }

    fn extension_formula(value: &str) -> Vec<u8> {
        format!(
            "<extLst xmlns:x=\"urn:generic\"><ext><x:f><![CDATA[{value}]]></x:f></ext></extLst>"
        )
        .into_bytes()
    }

    #[test]
    fn extension_formula_intersections_require_identity_axes() {
        let identity = SheetAxes::default();
        let mut moved = SheetAxes::default();
        moved.rows.insert(3, 1);
        let source = extension_formula("SUM($E$5 (E:E))");
        assert!(unchanged(&source, &identity));
        assert!(!unchanged(&source, &moved));
        assert!(!super::unchanged(&source, Some(&identity), &[], true));
    }

    #[test]
    fn extension_formula_defined_names_require_identity_axes() {
        let identity = SheetAxes::default();
        for names in [Some(&["AB1"][..]), None] {
            let source = extension_formula("ab1");
            assert!(super::unchanged_with_defined_names(
                &source,
                Some(&identity),
                &[],
                false,
                names
            ));
            assert!(!super::unchanged_with_defined_names(
                &source,
                Some(&identity),
                &[],
                true,
                names
            ));
        }
    }

    #[test]
    fn extension_formula_sheet_names_need_unique_save_origins() {
        let identity = SheetAxes::default();
        for sheets in [
            vec![("OldData", None), ("View", None), ("Data", None)],
            vec![("OldData", None), ("View", None), ("Other", None)],
            vec![("Data", Some(&identity)), ("data", Some(&identity))],
        ] {
            assert_eq!(
                formula("Data!E5", Some(&identity), &sheets, Some(&[])),
                ReferenceStatus::Unresolvable
            );
        }
    }

    #[test]
    fn extension_formula_unmoved_function_reference_allows_borrowing() {
        let identity = SheetAxes::default();
        let mut moved = SheetAxes::default();
        moved.rows.insert(3, 1);
        assert!(super::unchanged(
            &extension_formula("SUM(E5)"),
            Some(&identity),
            &[("Sheet1", Some(&moved)), ("Sheet2", Some(&identity))],
            true,
        ));
    }

    #[test]
    fn extension_formulas_follow_named_sheet_axes() {
        let mut current = SheetAxes::default();
        current.rows.insert(3, 1);
        let mut other = SheetAxes::default();
        other.cols.insert(3, 1);
        let identity = SheetAxes::default();
        let sheets = [
            ("Sheet1", Some(&current)),
            ("Sheet2", Some(&other)),
            ("Sheet 3's", Some(&identity)),
        ];
        for (value, stable) in [
            ("E5:E6", false),
            ("$E$5:$E$6", false),
            ("SUM(E1:E2)", true),
            ("Sheet1!E5:E6", false),
            ("'Sheet1'!$E$1:$E$2", true),
            ("sheet2!$E$1:$E$2", false),
            ("Sheet2!A5:B6", true),
            ("'Sheet 3''s'!$E$5:$E$6", true),
            ("$1:$2", true),
            ("$4:$5", false),
            ("$D:$E", true),
            ("Sheet2!$D:$E", false),
            ("Sheet2!$1:$5", true),
            ("Sheet2!$A : $B", true),
            ("Sheet2!A1:Sheet2!B2", true),
            ("'Sheet2'!A1:sheet2!B2", true),
            ("SUM(Sheet2!A1:B2,E5:E6)", false),
        ] {
            assert_eq!(
                super::unchanged(&extension_formula(value), Some(&current), &sheets, true),
                stable,
                "{value}"
            );
        }
    }

    #[test]
    fn extension_formula_unknown_references_require_identity_axes() {
        let identity = SheetAxes::default();
        let mut moved = SheetAxes::default();
        moved.rows.insert(3, 1);
        for value in [
            "Sheet2!A1",
            "'Missing Sheet'!A1",
            "Table1[Column1]",
            "[1]Sheet1!A1",
            "'[Book.xlsx]Sheet1'!A1",
            "Rate",
            "SUM(Rate,A1)",
            "NaN",
            "Sheet1:Sheet2!A1",
        ] {
            let source = extension_formula(value);
            assert!(
                super::unchanged(&source, Some(&identity), &[], false),
                "{value}"
            );
            assert!(
                !super::unchanged(
                    &source,
                    Some(&identity),
                    &[("Sheet1", Some(&moved)), ("Sheet2", None)],
                    true,
                ),
                "{value}"
            );
        }
        let source = extension_formula("A1");
        assert!(!super::unchanged(
            &source,
            None,
            &[("Sheet1", Some(&moved))],
            true
        ));
    }

    #[test]
    fn extension_formula_range_interiors_must_keep_their_indices() {
        for rows in [false, true] {
            let mut axes = SheetAxes::default();
            let axis = if rows { &mut axes.rows } else { &mut axes.cols };
            axis.delete(2, 1);
            axis.insert(2, 1);
            assert_eq!(axis.current(0), Some(0));
            assert_eq!(axis.current(4), Some(4));
            for value in ["A1:E5", "$E$5:$A$1", if rows { "1:5" } else { "A:E" }] {
                assert!(!unchanged(&extension_formula(value), &axes), "{value}");
            }
        }
    }

    #[test]
    fn extension_formula_constants_functions_and_strings_allow_borrowing() {
        let mut axes = SheetAxes::default();
        axes.rows.insert(3, 1);
        axes.cols.insert(3, 1);
        for value in [
            "",
            "1+2",
            "1E3+2.5e-4+1E+3",
            "IF(TRUE,SQRT(100),FALSE)",
            "_xlfn.FOO(1)",
            r#""A1"&"Sheet2!E5:E6"&"Table1[Column1]"&"A1""E5""#,
            "IFERROR(#REF!,#N/A)",
        ] {
            assert!(unchanged(&extension_formula(value), &axes), "{value}");
            assert!(
                super::unchanged(
                    &extension_formula(value),
                    None,
                    &[("Sheet1", Some(&axes))],
                    true,
                ),
                "{value}"
            );
        }
    }

    #[test]
    fn axis_edits_beyond_authored_coordinates_allow_borrowing() {
        let mut axes = SheetAxes::default();
        axes.rows.insert(100, 1);
        axes.cols.insert(100, 1);
        for source in [
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c r="A1"/></row></sheetData></worksheet>"#,
            r#"<sheetData><row outlineLevel="1"><c r="A1"><v>1</v></c></row></sheetData>"#,
            r#"<sheetData><row><c/><c/></row></sheetData>"#,
            "<sheetData>\n<row outlineLevel=\"1\">\n<c r=\"A1\"><v>1</v></c>\n</row>\n</sheetData>",
            r#"<sheetData><row r="5" outlineLevel="1"/></sheetData>"#,
            r#"<cols><col min="2" max="5" outlineLevel="1"/></cols>"#,
            r#"<mergeCells><mergeCell ref="B2:E5"/></mergeCells>"#,
            r#"<hyperlinks><hyperlink ref="B2:E5" location="Sheet1!A1"/></hyperlinks>"#,
            r#"<sheetData><row r="2"><c r="B2"><f t="array" ref="B2:E5">1</f><v>1</v></c></row></sheetData>"#,
            r#"<sheetData><row r="2"><c r="B2"><f t="shared" si="0" ref="B2:E5">1</f><v>1</v></c></row></sheetData>"#,
            r#"<sheetViews><sheetView><pane state="frozen" xSplit="2" ySplit="2" topLeftCell="C3"/><selection activeCell="E5" sqref="B2 E5"/></sheetView></sheetViews>"#,
            r#"<dataValidations><dataValidation sqref="B2:E5"/></dataValidations>"#,
            r#"<extLst><ext><sqref>B2:E5</sqref></ext></extLst>"#,
            r#"<f t="dataTable" r1="B2" r2="E5"/>"#,
            r#"<rowBreaks><brk id="4" min="0" max="4"/></rowBreaks>"#,
            r#"<rowBreaks><brk id="4" max="4"/></rowBreaks>"#,
            r#"<colBreaks><brk id="4" min="0" max="4"/></colBreaks>"#,
        ] {
            assert!(unchanged(source.as_bytes(), &axes), "{source}");
        }
    }

    #[test]
    fn moved_source_only_coordinates_refuse_borrowing() {
        for (source, row) in [
            (
                r#"<sheetData><row r="5" outlineLevel="1"/></sheetData>"#,
                true,
            ),
            (
                r#"<cols><col min="2" max="5" outlineLevel="1"/></cols>"#,
                false,
            ),
            (r#"<sheetData><row r="1" spans="2:5"/></sheetData>"#, false),
            (r#"<mergeCells><mergeCell ref="B2:E5"/></mergeCells>"#, true),
            (
                r#"<hyperlinks><hyperlink ref="B2:E5"/></hyperlinks>"#,
                false,
            ),
            (
                r#"<sheetData><row r="2"><c r="B2"><f t="array" ref="B2:E5">1</f></c></row></sheetData>"#,
                true,
            ),
            (
                r#"<sheetData><row r="2"><c r="B2"><f t="shared" si="0" ref="B2:E5">1</f></c></row></sheetData>"#,
                false,
            ),
            (r#"<pane state="frozen" ySplit="3"/>"#, true),
            (r#"<pane state="frozen" ySplit="3.0"/>"#, true),
            (r#"<selection activeCell="E5" sqref="A1"/>"#, false),
            (r#"<autoFilter ref="B2:E5"/>"#, true),
            (r#"<dataValidation sqref="B2:E5"/>"#, false),
            (r#"<f t="dataTable" r1="B2" r2="E5"/>"#, true),
            (r#"<extLst><ext><sqref>B2:E5</sqref></ext></extLst>"#, true),
            (
                r#"<extLst><ext><marker position="E5"/><ref>E5</ref></ext></extLst>"#,
                false,
            ),
            (
                r#"<extLst><ext><anchor><row>4</row><col>4</col></anchor></ext></extLst>"#,
                false,
            ),
        ] {
            let mut axes = SheetAxes::default();
            if row {
                axes.rows.insert(3, 1);
            } else {
                axes.cols.insert(3, 1);
            }
            assert!(!unchanged(source.as_bytes(), &axes), "{source}");
        }
    }

    #[test]
    fn deleted_and_reinserted_range_interiors_refuse_borrowing() {
        for rows in [false, true] {
            let mut axes = SheetAxes::default();
            let axis = if rows { &mut axes.rows } else { &mut axes.cols };
            axis.delete(2, 1);
            axis.insert(2, 1);
            assert_eq!(axis.current(0), Some(0));
            assert_eq!(axis.current(4), Some(4));
            assert_eq!(axis.current(2), None);
            for source in [
                r#"<mergeCell ref="A1:E5"/>"#,
                r#"<f t="shared" ref="A1:E5"/>"#,
                r#"<f t="array" ref="A1:E5"/>"#,
                r#"<hyperlink ref="A1:E5"/>"#,
                r#"<selection sqref="A1:E5"/>"#,
                r#"<extLst><ext><sqref><![CDATA[A1:E5]]></sqref></ext></extLst>"#,
                r#"<extLst><ext><sqref>&#65;1:&#69;5</sqref></ext></extLst>"#,
            ] {
                assert!(
                    !unchanged(source.as_bytes(), &axes),
                    "{source}, rows={rows}"
                );
            }
            let breaks = if rows {
                r#"<colBreaks><brk id="4" min="0" max="4"/></colBreaks>"#
            } else {
                r#"<rowBreaks><brk id="4" min="0" max="4"/></rowBreaks>"#
            };
            assert!(!unchanged(breaks.as_bytes(), &axes));
        }
    }

    #[test]
    fn deleted_and_reinserted_unauthored_gaps_allow_borrowing() {
        let source = r#"<sheetData><row r="1" outlineLevel="1"><c r="A1"/></row><row r="5" outlineLevel="2"><c r="E5"/></row></sheetData>"#;
        let mut axes = SheetAxes::default();
        axes.rows.delete(2, 1);
        axes.rows.insert(2, 1);
        axes.cols.delete(2, 1);
        axes.cols.insert(2, 1);
        assert!(unchanged(source.as_bytes(), &axes));
    }

    #[test]
    fn unreadable_positioned_markup_refuses_borrowing() {
        let mut axes = SheetAxes::default();
        axes.rows.insert(100, 1);
        for source in [
            r#"<row r="invalid"/>"#,
            r#"<mergeCell ref="invalid"/>"#,
            r#"<selection sqref="A1 invalid"/>"#,
            r#"<cols><col min="2"/></cols>"#,
        ] {
            assert!(!unchanged(source.as_bytes(), &axes), "{source}");
        }
    }

    #[test]
    fn unknown_attributes_and_extension_content_allow_borrowing() {
        let mut axes = SheetAxes::default();
        axes.rows.insert(3, 1);
        axes.cols.insert(3, 1);
        for source in [
            r#"<sheetPr><tabColor rgb="FF112233"/></sheetPr><sheetData><row outlineLevel="1"><c><v>1</v></c></row></sheetData>"#,
            r#"<sheetData><row custom="E5"><c custom="E5"><v>1</v></c></row></sheetData>"#,
            r#"<extLst><ext><marker position="E5" r="E5" ref="E5" sqref="E5" activeCell="E5" topLeftCell="E5" r1="E5" r2="E5" row="4" col="4" column="4"/><label>E5</label><value>4</value></ext></extLst>"#,
            r#"<extLst><ext><marker><row>4</row><col>4</col><column>4</column></marker><brk id="4" min="4" max="4"/></ext></extLst>"#,
        ] {
            assert!(unchanged(source.as_bytes(), &axes), "{source}");
        }
    }

    #[test]
    fn named_coordinate_attributes_and_extension_text_refuse_moved_axes() {
        for source in [
            r#"<dimension ref="B2:E5"/>"#,
            r#"<sheetView topLeftCell="E5"/>"#,
            r#"<conditionalFormatting sqref="B2:E5"/>"#,
            r#"<sortState ref="B2:E5"><sortCondition ref="E5"/></sortState>"#,
            r#"<protectedRange sqref="B2:E5"/>"#,
            r#"<ignoredError sqref="B2:E5"/>"#,
            r#"<cellWatch r="E5"/>"#,
            r#"<inputCells r="E5"/>"#,
            r#"<extLst xmlns:x="urn:generic"><ext><x:ref><![CDATA[B2:E5]]></x:ref></ext></extLst>"#,
            r#"<dataValidation xmlns:x="urn:generic" x:sqref="B2:E5"/>"#,
        ] {
            let mut axes = SheetAxes::default();
            axes.rows.insert(100, 1);
            axes.cols.insert(100, 1);
            assert!(unchanged(source.as_bytes(), &axes), "{source}");
            axes.rows.insert(3, 1);
            assert!(!unchanged(source.as_bytes(), &axes), "{source}");
        }
    }

    #[test]
    fn implicit_rows_and_cells_follow_reader_cursors() {
        for source in [
            r#"<sheetData><row/><row outlineLevel="1"/></sheetData>"#,
            r#"<sheetData><row r="5"/><row outlineLevel="1"/></sheetData>"#,
            r#"<sheetData><row r="5"><c r="C5"/><c/><c r="A5"/><c/></row><row><c/></row></sheetData>"#,
        ] {
            let mut axes = SheetAxes::default();
            axes.rows.insert(5, 1);
            axes.cols.insert(4, 1);
            assert!(unchanged(source.as_bytes(), &axes), "{source}");
            axes.rows.delete(0, 1);
            axes.rows.insert(0, 1);
            assert!(!unchanged(source.as_bytes(), &axes), "{source}");
        }
        let source = br#"<sheetData><row r="5"><c r="C5"/><c/></row></sheetData>"#;
        let mut axes = SheetAxes::default();
        axes.cols.insert(3, 1);
        assert!(!unchanged(source, &axes));
    }

    #[test]
    fn explicit_rows_use_reader_row_zero_fallback() {
        let mut axes = SheetAxes::default();
        axes.rows.insert(100, 1);
        for (source, expected) in [
            (
                r#"<sheetData><row/><row outlineLevel="1"/></sheetData>"#,
                r#"<sheetData><row r="1"/><row outlineLevel="1" r="1"/></sheetData>"#,
            ),
            (
                r#"<sheetData><row r="5"/><row outlineLevel="1"/></sheetData>"#,
                r#"<sheetData><row r="5"/><row outlineLevel="1" r="1"/></sheetData>"#,
            ),
            (
                r#"<sheetData><row r="5"></row><row><c/><c/></row></sheetData>"#,
                r#"<sheetData><row r="5"></row><row r="1"><c/><c/></row></sheetData>"#,
            ),
        ] {
            let explicit = explicit_rows(source.as_bytes(), &axes).unwrap().unwrap();
            assert_eq!(explicit, expected.as_bytes());
            assert!(unchanged(source.as_bytes(), &axes));
            assert!(unchanged(&explicit, &axes));
            let mut moved = axes.clone();
            moved.rows.delete(0, 1);
            moved.rows.insert(0, 1);
            assert!(!unchanged(source.as_bytes(), &moved));
            assert!(!unchanged(&explicit, &moved));
        }
    }
}
