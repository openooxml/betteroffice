use std::ops::Range;

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, Writer};
use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::{CellRange, CellRef};

use crate::axis::{AxisMap, SheetAxes};
use crate::package::attributes;
use crate::xml::{attr, resolve_entity, xml_err};
use crate::{MAX_DEPTH, ParseError};

pub(super) fn unchanged(source: &[u8], axes: &SheetAxes) -> bool {
    axes.is_identity() || inspect(source, axes).unwrap_or(false)
}

fn fixed(axis: &AxisMap, range: Range<u32>) -> bool {
    range.is_empty() || axis.current_ranges(range.clone()) == [range]
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

fn row_index(element: &BytesStart<'_>) -> Result<Option<u32>, ParseError> {
    Ok(match attr(element, b"r")? {
        Some(value) => value
            .trim()
            .parse::<u32>()
            .ok()
            .filter(|index| (1..=MAX_ROWS).contains(index))
            .map(|index| index - 1),
        None => Some(0),
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

fn inspect(source: &[u8], axes: &SheetAxes) -> Result<bool, ParseError> {
    let mut reader = Reader::from_reader(source);
    let mut parents: Vec<Vec<u8>> = Vec::new();
    let mut row = 0;
    let mut col = 0;
    let mut previous_row: Option<u32> = None;
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
                    let Some(index) = row_index(element)? else {
                        return Ok(false);
                    };
                    row = if attr(element, b"r")?.is_none() {
                        previous_row.map_or(0, |previous| previous + 1)
                    } else {
                        index
                    };
                    previous_row = Some(row);
                    col = 0;
                    if axes.rows.current(row) != Some(row) {
                        return Ok(false);
                    }
                }
                if source_cell {
                    let at = match attr(element, b"r")? {
                        Some(value) => CellRef::parse_a1(&value).ok(),
                        None => Some(CellRef::new(row, col)),
                    };
                    let Some(at) = at else {
                        return Ok(false);
                    };
                    col = at.col + 1;
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
                if name == b"brk" {
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
                    let stable = match key {
                        "r" if source_row || source_cell => true,
                        "r" | "ref" | "sqref" | "activeCell" | "topLeftCell" | "r1" | "r2" => {
                            references(value, axes)
                        }
                        "spans" if source_row => value.split_whitespace().all(|span| {
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
                        "xSplit" | "ySplit" if name == b"pane" => {
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
                        "row" | "col" | "column" => value.parse::<u32>().is_ok_and(|index| {
                            let axis = if key == "row" { &axes.rows } else { &axes.cols };
                            axis.current(index) == Some(index)
                        }),
                        "id" | "min" | "max" if name == b"brk" => {
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
                        _ => {
                            let mut values = value.split_whitespace().peekable();
                            values.peek().is_none()
                                || !values.all(|value| CellRange::parse_a1(value).is_ok())
                                || references(value, axes)
                        }
                    };
                    if !stable {
                        return Ok(false);
                    }
                }
                if matches!(event, Event::Start(_)) {
                    if name == b"sqref"
                        || (name == b"row" && !source_row)
                        || (name == b"col" && parent != Some(b"cols".as_slice()))
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
            Event::End(_) => {
                if let Some((name, value)) = positioned_text.take()
                    && !text_coordinates(&name, &value, axes)
                {
                    return Ok(false);
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
                    out.get_mut().extend_from_slice(&source[cursor..before]);
                    let mut element = element.clone();
                    element.push_attribute(("r", "1"));
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
                r#"<extLst><ext><marker position="E5"/></ext></extLst>"#,
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
}
