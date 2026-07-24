//! spreadsheetml -> `xlsx_model::Workbook`. streaming; nothing is sized from a
//! file-supplied `count`/`dimension`, cells and shared strings are capped.

use std::collections::BTreeMap;

use quick_xml::events::Event;
use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::{
    Cell, CellRange, CellRef, CellValue, DateSystem, DefinedName, ErrorValue, FreezePane,
    Hyperlink, Sheet, SheetId, Workbook,
};

use crate::styles::parse_stylesheet;
use crate::xml::{
    attr, collect_text, find_part, local_name, next_event, reader, resolve_part_path,
};
use crate::{MAX_CELLS, MAX_DEFINED_NAMES, MAX_HYPERLINKS, MAX_SHARED_STRINGS, ParseError};

/// parse a full workbook from opc parts, resolving sheets through the
/// workbook relationships.
pub fn parse_workbook(parts: &[(String, Vec<u8>)]) -> Result<Workbook, ParseError> {
    let wb_xml = find_part(parts, "xl/workbook.xml")
        .ok_or_else(|| ParseError::MissingPart("xl/workbook.xml".into()))?;
    let meta = parse_workbook_xml(wb_xml)?;

    let rels = find_part(parts, "xl/_rels/workbook.xml.rels")
        .map(parse_rels)
        .transpose()?
        .unwrap_or_default();

    let shared_strings = match find_part(parts, "xl/sharedStrings.xml") {
        Some(bytes) => parse_shared_strings(bytes)?,
        None => Vec::new(),
    };

    let wb_rels = find_part(parts, "xl/_rels/workbook.xml.rels");
    let styles_bytes = typed_part(parts, wb_rels, "styles", "xl/styles.xml")?;
    let theme_bytes = typed_part(parts, wb_rels, "theme", "xl/theme/theme1.xml")?;
    let styles = parse_stylesheet(styles_bytes, theme_bytes)?;

    let mut sheets = Vec::with_capacity(meta.sheets.len());
    for (idx, entry) in meta.sheets.iter().enumerate() {
        let path = worksheet_path(&rels, entry.rid.as_deref(), idx).ok_or_else(|| {
            ParseError::Malformed(format!("no target for sheet {:?}", entry.name))
        })?;
        let bytes = find_part(parts, &path).ok_or_else(|| ParseError::MissingPart(path.clone()))?;
        let sheet_rels = find_part(parts, &relationship_part_path(&path))
            .map(parse_rels)
            .transpose()?
            .unwrap_or_default();
        sheets.push(parse_worksheet(
            &entry.name,
            bytes,
            &shared_strings,
            &sheet_rels,
        )?);
    }

    Ok(Workbook {
        sheets,
        date_system: meta.date_system,
        defined_names: meta.defined_names,
        shared_strings,
        styles,
    })
}

/// resolve an optional part by relationship type suffix, falling back to
/// excel's conventional path when the rels are absent or lack the type.
fn typed_part<'a>(
    parts: &'a [(String, Vec<u8>)],
    wb_rels: Option<&[u8]>,
    type_suffix: &str,
    fallback: &str,
) -> Result<Option<&'a [u8]>, ParseError> {
    if let Some(rels) = wb_rels
        && let Some(target) = rel_target_by_type(rels, type_suffix)?
    {
        let path = resolve_part_path("xl", &target);
        return Ok(find_part(parts, &path));
    }
    Ok(find_part(parts, fallback))
}

/// find the `Target` of the first `Relationship` whose `Type` ends with
/// `/{type_suffix}`.
fn rel_target_by_type(data: &[u8], type_suffix: &str) -> Result<Option<String>, ParseError> {
    let mut reader = reader(data);
    let mut buf = Vec::new();
    let mut depth = 0;
    let needle = format!("/{type_suffix}");

    loop {
        match next_event(&mut reader, &mut buf, &mut depth)? {
            Event::Start(e) if local_name(&e) == b"Relationship" => {
                if let (Some(ty), Some(target)) = (attr(&e, b"Type")?, attr(&e, b"Target")?)
                    && ty.ends_with(&needle)
                {
                    return Ok(Some(target));
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(None)
}

struct SheetEntry {
    name: String,
    rid: Option<String>,
}

struct WorkbookMeta {
    date_system: DateSystem,
    sheets: Vec<SheetEntry>,
    defined_names: Vec<DefinedName>,
}

/// read sheet order/names, the r:id linking each to a worksheet part, and the
/// 1900/1904 date epoch flag.
fn parse_workbook_xml(data: &[u8]) -> Result<WorkbookMeta, ParseError> {
    let mut reader = reader(data);
    let mut buf = Vec::new();
    let mut depth = 0;
    let mut date_system = DateSystem::V1900;
    let mut sheets = Vec::new();
    let mut defined_names = Vec::new();

    loop {
        match next_event(&mut reader, &mut buf, &mut depth)? {
            Event::Start(e) => match local_name(&e).as_slice() {
                b"workbookPr" => {
                    if let Some(v) = attr(&e, b"date1904")?
                        && is_truthy(&v)
                    {
                        date_system = DateSystem::V1904;
                    }
                }
                b"sheet" => {
                    let name = attr(&e, b"name")?.unwrap_or_default();
                    let rid = attr(&e, b"id")?;
                    sheets.push(SheetEntry { name, rid });
                }
                b"definedName" => {
                    if defined_names.len() >= MAX_DEFINED_NAMES {
                        return Err(ParseError::TooManyDefinedNames);
                    }
                    let name = attr(&e, b"name")?.unwrap_or_default();
                    let local_sheet = attr(&e, b"localSheetId")?
                        .and_then(|value| value.parse::<u32>().ok())
                        .map(SheetId);
                    let hidden = attr(&e, b"hidden")?.is_some_and(|value| is_truthy(&value));
                    let formula = collect_text(&mut reader, &mut buf, &mut depth)?;
                    defined_names.push(DefinedName {
                        name,
                        formula,
                        local_sheet,
                        hidden,
                    });
                }
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(WorkbookMeta {
        date_system,
        sheets,
        defined_names,
    })
}

#[derive(Clone)]
struct Relationship {
    target: String,
    kind: Option<String>,
    external: bool,
}

/// map relationship id -> relationship metadata from a `.rels` part.
fn parse_rels(data: &[u8]) -> Result<BTreeMap<String, Relationship>, ParseError> {
    let mut reader = reader(data);
    let mut buf = Vec::new();
    let mut depth = 0;
    let mut map = BTreeMap::new();

    loop {
        match next_event(&mut reader, &mut buf, &mut depth)? {
            Event::Start(e) if local_name(&e) == b"Relationship" => {
                if let (Some(id), Some(target)) = (attr(&e, b"Id")?, attr(&e, b"Target")?) {
                    let kind = attr(&e, b"Type")?;
                    let external = attr(&e, b"TargetMode")?
                        .is_some_and(|mode| mode.eq_ignore_ascii_case("external"));
                    map.insert(
                        id,
                        Relationship {
                            target,
                            kind,
                            external,
                        },
                    );
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(map)
}

/// pick the worksheet part path: the relationship target, else the
/// conventional positional name.
fn worksheet_path(
    rels: &BTreeMap<String, Relationship>,
    rid: Option<&str>,
    idx: usize,
) -> Option<String> {
    if let Some(relationship) = rid.and_then(|r| rels.get(r))
        && !relationship.external
    {
        return Some(resolve_part_path("xl", &relationship.target));
    }
    Some(format!("xl/worksheets/sheet{}.xml", idx + 1))
}

fn relationship_part_path(path: &str) -> String {
    match path.rsplit_once('/') {
        Some((directory, file)) => format!("{directory}/_rels/{file}.rels"),
        None => format!("_rels/{path}.rels"),
    }
}

/// parse the shared string table, flattening rich runs to plain text.
fn parse_shared_strings(data: &[u8]) -> Result<Vec<String>, ParseError> {
    let mut reader = reader(data);
    let mut buf = Vec::new();
    let mut depth = 0;
    let mut strings = Vec::new();

    loop {
        match next_event(&mut reader, &mut buf, &mut depth)? {
            Event::Start(e) if local_name(&e) == b"si" => {
                if strings.len() >= MAX_SHARED_STRINGS {
                    return Err(ParseError::TooManyStrings);
                }
                strings.push(collect_text(&mut reader, &mut buf, &mut depth)?);
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(strings)
}

/// in-progress cell state accumulated between a `<c>` start and its end.
#[derive(Default)]
struct CellBuild {
    addr: Option<CellRef>,
    ty: Option<String>,
    style: Option<u32>,
    value_text: Option<String>,
    inline_text: Option<String>,
    formula: Option<String>,
}

/// parse one worksheet into a `Sheet`: cells (values, cached formulas, types),
/// merges, and column/row sizing. `shared` resolves `t="s"` indices.
fn parse_worksheet(
    name: &str,
    data: &[u8],
    shared: &[String],
    relationships: &BTreeMap<String, Relationship>,
) -> Result<Sheet, ParseError> {
    let mut reader = reader(data);
    let mut buf = Vec::new();
    let mut depth = 0;
    let mut sheet = Sheet::new(name);
    let mut cur_row: Option<u32> = None;
    let mut col_cursor: u32 = 0;
    let mut cur: Option<CellBuild> = None;
    let mut cell_count: u64 = 0;
    let mut hyperlink_count: usize = 0;

    loop {
        match next_event(&mut reader, &mut buf, &mut depth)? {
            Event::Start(e) => match local_name(&e).as_slice() {
                b"row" => {
                    let row = match attr(&e, b"r")? {
                        Some(v) => parse_index(&v, MAX_ROWS)?,
                        None => cur_row.map_or(0, |r| r + 1),
                    };
                    cur_row = Some(row);
                    col_cursor = 0;
                    if let Some(h) = attr(&e, b"ht")?.and_then(|v| v.parse::<f64>().ok()) {
                        sheet.row_heights.insert(row, h);
                    }
                }
                b"c" => {
                    cell_count += 1;
                    if cell_count > MAX_CELLS {
                        return Err(ParseError::TooManyCells);
                    }
                    let addr = match attr(&e, b"r")? {
                        Some(v) => CellRef::parse_a1(&v)
                            .map_err(|_| ParseError::Malformed(format!("bad cell ref {v:?}")))?,
                        None => CellRef::new(cur_row.unwrap_or(0), col_cursor),
                    };
                    col_cursor = addr.col;
                    let style = attr(&e, b"s")?.and_then(|v| v.parse::<u32>().ok());
                    cur = Some(CellBuild {
                        addr: Some(addr),
                        ty: attr(&e, b"t")?,
                        style,
                        ..CellBuild::default()
                    });
                }
                b"v" => {
                    let text = collect_text(&mut reader, &mut buf, &mut depth)?;
                    if let Some(c) = cur.as_mut() {
                        c.value_text = Some(text);
                    }
                }
                b"f" => {
                    let text = collect_text(&mut reader, &mut buf, &mut depth)?;
                    if let Some(c) = cur.as_mut() {
                        c.formula = Some(text);
                    }
                }
                b"is" => {
                    let text = collect_text(&mut reader, &mut buf, &mut depth)?;
                    if let Some(c) = cur.as_mut() {
                        c.inline_text = Some(text);
                    }
                }
                b"mergeCell" => {
                    if let Some(r) = attr(&e, b"ref")? {
                        let range = CellRange::parse_a1(&r)
                            .map_err(|_| ParseError::Malformed(format!("bad merge ref {r:?}")))?;
                        sheet.merges.push(range);
                    }
                }
                b"pane" => {
                    if sheet.freeze_pane.is_none() {
                        sheet.freeze_pane = parse_freeze_pane(&e)?;
                    }
                }
                b"hyperlink" => {
                    hyperlink_count += 1;
                    if hyperlink_count > MAX_HYPERLINKS {
                        return Err(ParseError::TooManyHyperlinks);
                    }
                    if let Some(link) = parse_hyperlink(&e, relationships)? {
                        sheet.hyperlinks.push(link);
                    }
                }
                b"col" => parse_col(&e, &mut sheet)?,
                _ => {}
            },
            Event::End(e) => {
                let name = e.name();
                match name.local_name().as_ref() {
                    b"c" => {
                        if let Some(c) = cur.take() {
                            finalize_cell(c, shared, &mut sheet)?;
                        }
                        col_cursor += 1;
                    }
                    b"row" => cur_row = None,
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    normalize_merges(&mut sheet.merges);
    Ok(sheet)
}

fn parse_hyperlink(
    element: &quick_xml::events::BytesStart,
    relationships: &BTreeMap<String, Relationship>,
) -> Result<Option<Hyperlink>, ParseError> {
    let Some(reference) = attr(element, b"ref")? else {
        return Ok(None);
    };
    let range = CellRange::parse_a1(&reference)
        .map_err(|_| ParseError::Malformed(format!("bad hyperlink ref {reference:?}")))?;
    let external_target = attr(element, b"id")?
        .and_then(|id| relationships.get(&id))
        .filter(|relationship| {
            relationship.external
                && relationship
                    .kind
                    .as_deref()
                    .is_some_and(|kind| kind.ends_with("/hyperlink"))
        })
        .map(|relationship| relationship.target.clone())
        .filter(|target| !target.is_empty());
    let location = attr(element, b"location")?.filter(|location| !location.is_empty());
    if external_target.is_none() && location.is_none() {
        return Ok(None);
    }
    Ok(Some(Hyperlink {
        range,
        external_target,
        location,
        tooltip: attr(element, b"tooltip")?,
        display: attr(element, b"display")?,
    }))
}

fn parse_freeze_pane(
    element: &quick_xml::events::BytesStart,
) -> Result<Option<FreezePane>, ParseError> {
    let state = attr(element, b"state")?;
    if !matches!(state.as_deref(), Some("frozen" | "frozenSplit")) {
        return Ok(None);
    }
    let rows = frozen_count(attr(element, b"ySplit")?, MAX_ROWS);
    let cols = frozen_count(attr(element, b"xSplit")?, MAX_COLS);
    if rows == 0 && cols == 0 {
        return Ok(None);
    }
    let fallback = CellRef::new(
        rows.min(MAX_ROWS.saturating_sub(1)),
        cols.min(MAX_COLS.saturating_sub(1)),
    );
    let top_left = attr(element, b"topLeftCell")?
        .and_then(|value| CellRef::parse_a1(&value).ok())
        .unwrap_or(fallback);
    Ok(Some(FreezePane::new(rows, cols, top_left)))
}

fn frozen_count(value: Option<String>, limit: u32) -> u32 {
    value
        .and_then(|value| value.parse::<f64>().ok())
        .filter(|value| value.is_finite() && *value >= 0.0 && value.fract() == 0.0)
        .map(|value| value.min(f64::from(limit)) as u32)
        .unwrap_or(0)
}

fn normalize_merges(merges: &mut Vec<CellRange>) {
    let mut index = 0;
    while index < merges.len() {
        let range = merges[index];
        if merges[..index]
            .iter()
            .any(|kept| ranges_intersect(*kept, range))
        {
            merges.remove(index);
        } else {
            index += 1;
        }
    }
}

fn ranges_intersect(left: CellRange, right: CellRange) -> bool {
    left.start.row <= right.end.row
        && left.end.row >= right.start.row
        && left.start.col <= right.end.col
        && left.end.col >= right.start.col
}

/// apply a `<col>` width across its `[min, max]` span (clamped to sheet bounds).
/// widths are stored per-column since the model has no column-range concept.
fn parse_col(e: &quick_xml::events::BytesStart, sheet: &mut Sheet) -> Result<(), ParseError> {
    let width = match attr(e, b"width")?.and_then(|v| v.parse::<f64>().ok()) {
        Some(w) => w,
        None => return Ok(()),
    };
    let min = attr(e, b"min")?
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(1);
    let max = attr(e, b"max")?
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(min);
    let min = min.clamp(1, MAX_COLS);
    let max = max.clamp(min, MAX_COLS);
    for col in min..=max {
        sheet.col_widths.insert(col - 1, width);
    }
    Ok(())
}

/// turn accumulated cell state into a stored `Cell`, decoding the value per its
/// `t` type. an empty non-styled, non-formula cell is dropped.
fn finalize_cell(c: CellBuild, shared: &[String], sheet: &mut Sheet) -> Result<(), ParseError> {
    let addr = match c.addr {
        Some(a) => a,
        None => return Ok(()),
    };
    let value = match c.ty.as_deref() {
        Some("s") => {
            let idx = c
                .value_text
                .as_deref()
                .and_then(|v| v.trim().parse::<usize>().ok());
            match idx.and_then(|i| shared.get(i)) {
                Some(s) => CellValue::Text { value: s.clone() },
                None => CellValue::Empty,
            }
        }
        Some("inlineStr") => CellValue::Text {
            value: c.inline_text.unwrap_or_default(),
        },
        Some("str") => CellValue::Text {
            value: c.value_text.unwrap_or_default(),
        },
        Some("b") => CellValue::Bool {
            value: c.value_text.as_deref() == Some("1"),
        },
        Some("e") => match c.value_text.as_deref().and_then(error_from_str) {
            Some(err) => CellValue::Error { value: err },
            None => CellValue::Text {
                value: c.value_text.clone().unwrap_or_default(),
            },
        },
        _ => match c.value_text.as_deref() {
            Some(v) if !v.trim().is_empty() => {
                let n = v
                    .trim()
                    .parse::<f64>()
                    .map_err(|_| ParseError::Malformed(format!("bad number {v:?}")))?;
                CellValue::Number { value: n }
            }
            _ => CellValue::Empty,
        },
    };

    let cell = Cell {
        value,
        formula: c.formula,
        style: c.style,
    };
    if cell != Cell::default() {
        sheet.set_cell(addr, cell);
    }
    Ok(())
}

fn error_from_str(s: &str) -> Option<ErrorValue> {
    Some(match s {
        "#DIV/0!" => ErrorValue::Div0,
        "#N/A" => ErrorValue::NA,
        "#NAME?" => ErrorValue::Name,
        "#NULL!" => ErrorValue::Null,
        "#NUM!" => ErrorValue::Num,
        "#REF!" => ErrorValue::Ref,
        "#VALUE!" => ErrorValue::Value,
        "#SPILL!" => ErrorValue::Spill,
        _ => return None,
    })
}

fn is_truthy(v: &str) -> bool {
    matches!(v, "1" | "true" | "on")
}

/// parse a 1-based xml index (row number) into a 0-based id, bounds-checked.
fn parse_index(v: &str, max: u32) -> Result<u32, ParseError> {
    let n: u32 = v
        .parse()
        .map_err(|_| ParseError::Malformed(format!("bad index {v:?}")))?;
    if n == 0 || n > max {
        return Err(ParseError::Malformed(format!("index out of range {v:?}")));
    }
    Ok(n - 1)
}
