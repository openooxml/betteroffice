//! Lexical patching of a preserved worksheet's `<cols>` and `<sheetData>`.

use std::collections::{BTreeSet, HashMap, HashSet};
use std::iter::Peekable;
use std::ops::Range;

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::{Cell, CellRange, CellRef, Sheet, Workbook};

use crate::axis::SheetAxes;
use crate::package::{XmlAttribute, attributes, remove_attribute, set_attribute};
use crate::read::SharedStringCells;
use crate::write::{
    SharedStringPlan, fmt_num, fragment, shared_string_index, write_cell, write_col, write_cols,
    write_row,
};
use crate::xml::{attr, xml_err};
use crate::{MAX_DEPTH, ParseError};

/// One edited sheet against the source it was read from.
pub(crate) struct SheetPatch<'a> {
    pub(crate) sheet: &'a Sheet,
    pub(crate) original: &'a Sheet,
    pub(crate) axes: &'a SheetAxes,
    pub(crate) workbook: &'a Workbook,
    pub(crate) sst_index: &'a HashMap<&'a str, usize>,
    pub(crate) retained: &'a SharedStringCells,
    pub(crate) plan: Option<&'a SharedStringPlan>,
}

struct SourceElement {
    prefix: Vec<u8>,
    suffix: Vec<u8>,
    content: Range<usize>,
}

struct SourceRow {
    before: Range<usize>,
    span: Range<usize>,
    tag: Range<usize>,
    empty: bool,
    index: u32,
    cells: Vec<SourceCell>,
}

impl SourceRow {
    fn content_start(&self) -> usize {
        self.cells.last().map_or(self.tag.end, |cell| cell.span.end)
    }
}

struct SourceCell {
    before: Range<usize>,
    span: Range<usize>,
    tag: Range<usize>,
    empty: bool,
    at: CellRef,
    shared_string: bool,
    formula: Option<FormulaMarkup>,
}

/// The `<f>` attributes that tie a formula to where it sits.
struct FormulaMarkup {
    tag: Range<usize>,
    empty: bool,
    group: Option<u32>,
    reference: Option<CellRange>,
    /// carries `r1`/`r2` or something else this crate cannot move.
    positional: bool,
}

struct SourceCol {
    before: Range<usize>,
    span: Range<usize>,
    tag: Range<usize>,
    empty: bool,
    range: Range<u32>,
}

/// Shared-formula groups and array masters to rewrite whole.
struct DirtyFormulas {
    groups: HashSet<u32>,
    masters: HashSet<(u32, u32)>,
}

struct WidthPiece {
    range: Range<u32>,
    changed: Option<Option<f64>>,
}

impl SheetPatch<'_> {
    /// Patched `<sheetData>`; `None` when the source resists cell-by-cell patching.
    pub(crate) fn sheet_data(&self, source: &[u8]) -> Result<Option<Vec<u8>>, ParseError> {
        let Some((element, rows)) = scan_sheet_data(source)? else {
            return Ok(None);
        };
        let dirty = self.dirty_formulas(&rows);
        let mut out = Vec::with_capacity(source.len());
        out.extend_from_slice(&element.prefix);
        let mut source_rows = rows
            .iter()
            .filter_map(|row| {
                self.axes
                    .rows
                    .current(row.index)
                    .map(|current| (current, row))
            })
            .peekable();
        let mut cells = self.sheet.iter_cells().peekable();
        let mut heights = self.sheet.row_heights.keys().copied().peekable();
        loop {
            let next_source = source_rows.peek().map(|(current, _)| *current);
            let next_cell = cells.peek().map(|(at, _)| at.row);
            let next_height = heights.peek().copied();
            let Some(row) = [next_source, next_cell, next_height]
                .into_iter()
                .flatten()
                .min()
            else {
                break;
            };
            if next_height == Some(row) {
                heights.next();
            }
            if next_source == Some(row) {
                let (_, source_row) = source_rows.next().expect("peeked");
                self.emit_source_row(&mut out, source, source_row, row, &mut cells, &dirty)?;
            } else {
                self.emit_generated_row(&mut out, row, &mut cells)?;
            }
        }
        let tail = rows
            .last()
            .map_or(element.content.start, |row| row.span.end);
        out.extend_from_slice(&source[tail..element.content.end]);
        out.extend_from_slice(&element.suffix);
        Ok(Some(out))
    }

    /// Patched `<cols>`; `None` removes the element.
    pub(crate) fn cols(&self, source: Option<&[u8]>) -> Result<Option<Vec<u8>>, ParseError> {
        let Some(source) = source else {
            return self.generated_cols();
        };
        let Some((element, cols)) = scan_cols(source)? else {
            return self.generated_cols();
        };
        let mut entries: Vec<(u32, Vec<u8>, bool)> = Vec::new();
        let mut covered = BTreeSet::new();
        for col in &cols {
            let (name, attributes) = start_tag(&source[col.tag.clone()])?;
            let pieces = self
                .axes
                .cols
                .current_ranges(col.range.clone())
                .into_iter()
                .flat_map(|range| self.width_pieces(range))
                .collect::<Vec<_>>();
            let verbatim =
                pieces.len() == 1 && pieces[0].range == col.range && pieces[0].changed.is_none();
            for (index, piece) in pieces.iter().enumerate() {
                let mut bytes = Vec::new();
                if index == 0 {
                    bytes.extend_from_slice(&source[col.before.clone()]);
                }
                if verbatim {
                    bytes.extend_from_slice(&source[col.span.clone()]);
                } else {
                    let mut attributes = attributes.clone();
                    set_attribute(
                        &mut attributes,
                        "min",
                        "min",
                        (piece.range.start + 1).to_string(),
                    );
                    set_attribute(&mut attributes, "max", "max", piece.range.end.to_string());
                    if let Some(width) = piece.changed {
                        set_col_width(&mut attributes, width);
                    }
                    write_start_tag(&mut bytes, &name, &attributes, col.empty)?;
                    if !col.empty {
                        bytes.extend_from_slice(&source[col.tag.end..col.span.end]);
                    }
                }
                covered.extend(piece.range.clone());
                entries.push((piece.range.start, bytes, false));
            }
        }
        for (&col, &width) in &self.sheet.col_widths {
            if covered.contains(&col) {
                continue;
            }
            let mut writer = Writer::new(Vec::new());
            write_col(&mut writer, col, width).map_err(xml_err)?;
            let at = entries
                .iter()
                .position(|(min, _, generated)| !generated && *min > col)
                .unwrap_or(entries.len());
            entries.insert(at, (col, writer.into_inner(), true));
        }
        if entries.is_empty() {
            return Ok(None);
        }
        let mut out = element.prefix;
        for (_, bytes, _) in entries {
            out.extend_from_slice(&bytes);
        }
        let tail = cols
            .last()
            .map_or(element.content.start, |col| col.span.end);
        out.extend_from_slice(&source[tail..element.content.end]);
        out.extend_from_slice(&element.suffix);
        Ok(Some(out))
    }

    fn generated_cols(&self) -> Result<Option<Vec<u8>>, ParseError> {
        (!self.sheet.col_widths.is_empty())
            .then(|| fragment(|writer| write_cols(writer, self.sheet)))
            .transpose()
    }

    /// Splits a column span wherever the model width parted from the source.
    fn width_pieces(&self, range: Range<u32>) -> Vec<WidthPiece> {
        let mut pieces: Vec<WidthPiece> = Vec::new();
        for col in range {
            let model = self.sheet.col_widths.get(&col).copied();
            let original = self
                .axes
                .cols
                .source(col)
                .and_then(|source| self.original.col_widths.get(&source).copied());
            let changed = (model != original).then_some(model);
            match pieces.last_mut() {
                Some(piece) if piece.changed == changed => piece.range.end = col + 1,
                _ => pieces.push(WidthPiece {
                    range: col..col + 1,
                    changed,
                }),
            }
        }
        pieces
    }

    fn emit_generated_row<'c, I>(
        &self,
        out: &mut Vec<u8>,
        row: u32,
        cells: &mut Peekable<I>,
    ) -> Result<(), ParseError>
    where
        I: Iterator<Item = (CellRef, &'c Cell)>,
    {
        let mut writer = Writer::new(std::mem::take(out));
        write_row(
            &mut writer,
            self.sheet,
            row,
            cells,
            self.workbook,
            self.sst_index,
            self.retained,
            self.plan,
        )
        .map_err(xml_err)?;
        *out = writer.into_inner();
        Ok(())
    }

    fn emit_source_row<'c, I>(
        &self,
        out: &mut Vec<u8>,
        data: &[u8],
        source: &SourceRow,
        row: u32,
        cells: &mut Peekable<I>,
        dirty: &DirtyFormulas,
    ) -> Result<(), ParseError>
    where
        I: Iterator<Item = (CellRef, &'c Cell)>,
    {
        let mut body = Vec::new();
        let mut columns: Option<(u32, u32)> = None;
        let mut source_cells = source
            .cells
            .iter()
            .filter_map(|cell| {
                self.axes
                    .cols
                    .current(cell.at.col)
                    .map(|current| (current, cell))
            })
            .peekable();
        loop {
            let next_source = source_cells.peek().map(|(current, _)| *current);
            let next_model = cells
                .peek()
                .filter(|(at, _)| at.row == row)
                .map(|(at, _)| at.col);
            let Some(col) = [next_source, next_model].into_iter().flatten().min() else {
                break;
            };
            let source_cell =
                (next_source == Some(col)).then(|| source_cells.next().expect("peeked").1);
            let model_cell = (next_model == Some(col)).then(|| cells.next().expect("peeked").1);
            let at = CellRef::new(row, col);
            match source_cell {
                Some(source_cell) if self.verbatim(source_cell, at, dirty) => {
                    body.extend_from_slice(&data[source_cell.before.clone()]);
                    self.emit_source_cell(&mut body, data, source_cell, at)?;
                }
                Some(source_cell) => {
                    let Some(cell) = model_cell else {
                        continue;
                    };
                    body.extend_from_slice(&data[source_cell.before.clone()]);
                    self.emit_cell(&mut body, at, cell)?;
                }
                None => self.emit_cell(&mut body, at, model_cell.expect("one side is present"))?,
            }
            columns = Some(match columns {
                Some((min, max)) => (min, max.max(col + 1)),
                None => (col + 1, col + 1),
            });
        }

        let moved = source.index != row;
        let height = self.sheet.row_heights.get(&row).copied();
        let height_changed = height != self.original.row_heights.get(&source.index).copied();
        let (name, mut attributes) = start_tag(&data[source.tag.clone()])?;
        let mut rewrite = moved || height_changed;
        if moved {
            set_attribute(&mut attributes, "r", "r", (u64::from(row) + 1).to_string());
        }
        if height_changed {
            set_row_height(&mut attributes, height);
        }
        if let Some((min, max)) = columns
            && spans_exclude(&attributes, min, max)
        {
            set_attribute(&mut attributes, "spans", "spans", format!("{min}:{max}"));
            rewrite = true;
        }

        out.extend_from_slice(&data[source.before.clone()]);
        let empty = source.empty && body.is_empty();
        if rewrite || (source.empty && !empty) {
            write_start_tag(out, &name, &attributes, empty)?;
        } else {
            out.extend_from_slice(&data[source.tag.clone()]);
        }
        if empty {
            return Ok(());
        }
        out.extend_from_slice(&body);
        if source.empty {
            write_end_tag(out, &name)?;
        } else {
            out.extend_from_slice(&data[source.content_start()..source.span.end]);
        }
        Ok(())
    }

    /// Whether a source cell can stand for the model cell now at `at`.
    fn verbatim(&self, cell: &SourceCell, at: CellRef, dirty: &DirtyFormulas) -> bool {
        let model = self.sheet.cell(at);
        if model != self.original.cell(cell.at) {
            return false;
        }
        if cell.shared_string
            && let Some(model) = model
        {
            let Some(source) = self.retained.get(&(at.row, at.col)).copied() else {
                return false;
            };
            let written = shared_string_index(
                model,
                at,
                self.workbook,
                self.sst_index,
                self.retained,
                self.plan,
            );
            if written != Some(source) {
                return false;
            }
        }
        let Some(formula) = &cell.formula else {
            return true;
        };
        if cell.at != at && formula.positional {
            return false;
        }
        match formula.group {
            Some(group) => !dirty.groups.contains(&group),
            None => {
                formula.reference.is_none() || !dirty.masters.contains(&(cell.at.row, cell.at.col))
            }
        }
    }

    fn emit_source_cell(
        &self,
        out: &mut Vec<u8>,
        data: &[u8],
        cell: &SourceCell,
        at: CellRef,
    ) -> Result<(), ParseError> {
        if cell.at == at {
            out.extend_from_slice(&data[cell.span.clone()]);
            return Ok(());
        }
        let (name, mut attributes) = start_tag(&data[cell.tag.clone()])?;
        set_attribute(&mut attributes, "r", "r", at.to_a1());
        write_start_tag(out, &name, &attributes, cell.empty)?;
        let mut cursor = cell.tag.end;
        if let Some(formula) = &cell.formula
            && let Some(reference) = formula.reference
            && let Some(remapped) = self.remap_range(reference)
            && remapped != reference
        {
            out.extend_from_slice(&data[cursor..formula.tag.start]);
            let (name, mut attributes) = start_tag(&data[formula.tag.clone()])?;
            set_attribute(&mut attributes, "ref", "ref", remapped.to_a1());
            write_start_tag(out, &name, &attributes, formula.empty)?;
            cursor = formula.tag.end;
        }
        out.extend_from_slice(&data[cursor..cell.span.end]);
        Ok(())
    }

    fn emit_cell(&self, out: &mut Vec<u8>, at: CellRef, cell: &Cell) -> Result<(), ParseError> {
        let retained = shared_string_index(
            cell,
            at,
            self.workbook,
            self.sst_index,
            self.retained,
            self.plan,
        );
        let mut writer = Writer::new(std::mem::take(out));
        write_cell(
            &mut writer,
            at,
            cell,
            self.sst_index,
            retained,
            self.sheet.array_formula(at),
        )
        .map_err(xml_err)?;
        *out = writer.into_inner();
        Ok(())
    }

    fn dirty_formulas(&self, rows: &[SourceRow]) -> DirtyFormulas {
        let changed = self.changed_source_cells();
        let mut groups: HashMap<u32, (bool, Option<CellRange>)> = HashMap::new();
        let mut masters = Vec::new();
        for cell in rows.iter().flat_map(|row| &row.cells) {
            let Some(formula) = &cell.formula else {
                continue;
            };
            let key = (cell.at.row, cell.at.col);
            match formula.group {
                Some(group) => {
                    let entry = groups.entry(group).or_insert((false, None));
                    entry.0 |= changed.contains(&key);
                    if let Some(reference) = formula.reference
                        && entry.1.replace(reference).is_some()
                    {
                        entry.0 = true;
                    }
                }
                None => {
                    if let Some(reference) = formula.reference {
                        masters.push((key, reference));
                    }
                }
            }
        }
        DirtyFormulas {
            groups: groups
                .into_iter()
                .filter(|(_, (dirty, master))| {
                    *dirty || !master.is_some_and(|reference| self.moves_uniformly(reference))
                })
                .map(|(group, _)| group)
                .collect(),
            masters: masters
                .into_iter()
                .filter(|(key, reference)| {
                    changed.contains(key)
                        || !self.moves_uniformly(*reference)
                        || range_changed(*reference, &changed)
                })
                .map(|(key, _)| key)
                .collect(),
        }
    }

    /// Source addresses whose cell changed, moved, or is gone.
    fn changed_source_cells(&self) -> BTreeSet<(u32, u32)> {
        let mut changed = BTreeSet::new();
        for (at, cell) in self.original.iter_cells() {
            if self.mapped(at).and_then(|mapped| self.sheet.cell(mapped)) != Some(cell) {
                changed.insert((at.row, at.col));
            }
        }
        for (at, cell) in self.sheet.iter_cells() {
            if let Some(source) = self.inverse(at)
                && self.original.cell(source) != Some(cell)
            {
                changed.insert((source.row, source.col));
            }
        }
        changed
    }

    fn moves_uniformly(&self, reference: CellRange) -> bool {
        self.remap_range(reference).is_some_and(|remapped| {
            remapped.end.row - remapped.start.row == reference.end.row - reference.start.row
                && remapped.end.col - remapped.start.col == reference.end.col - reference.start.col
        })
    }

    fn remap_range(&self, reference: CellRange) -> Option<CellRange> {
        Some(CellRange::new(
            self.mapped(reference.start)?,
            self.mapped(reference.end)?,
        ))
    }

    fn mapped(&self, at: CellRef) -> Option<CellRef> {
        Some(CellRef::new(
            self.axes.rows.current(at.row)?,
            self.axes.cols.current(at.col)?,
        ))
    }

    fn inverse(&self, at: CellRef) -> Option<CellRef> {
        Some(CellRef::new(
            self.axes.rows.source(at.row)?,
            self.axes.cols.source(at.col)?,
        ))
    }
}

fn range_changed(reference: CellRange, changed: &BTreeSet<(u32, u32)>) -> bool {
    changed
        .range((reference.start.row, 0)..=(reference.end.row, u32::MAX))
        .any(|&(_, col)| col >= reference.start.col && col <= reference.end.col)
}

fn set_row_height(attributes: &mut Vec<XmlAttribute>, height: Option<f64>) {
    remove_attribute(attributes, "ht");
    remove_attribute(attributes, "customHeight");
    remove_attribute(attributes, "hidden");
    if let Some(height) = height {
        let value = fmt_num(height);
        let hidden = value == "0";
        set_attribute(attributes, "ht", "ht", value);
        set_attribute(attributes, "customHeight", "customHeight", "1".to_owned());
        if hidden {
            set_attribute(attributes, "hidden", "hidden", "1".to_owned());
        }
    }
}

fn set_col_width(attributes: &mut Vec<XmlAttribute>, width: Option<f64>) {
    remove_attribute(attributes, "width");
    remove_attribute(attributes, "customWidth");
    remove_attribute(attributes, "hidden");
    remove_attribute(attributes, "bestFit");
    if let Some(width) = width {
        let value = fmt_num(width);
        let hidden = value == "0";
        set_attribute(attributes, "width", "width", value);
        set_attribute(attributes, "customWidth", "customWidth", "1".to_owned());
        if hidden {
            set_attribute(attributes, "hidden", "hidden", "1".to_owned());
        }
    }
}

/// Whether an authored `spans` hint no longer covers the row's cells.
fn spans_exclude(attributes: &[XmlAttribute], min: u32, max: u32) -> bool {
    let Some(spans) = attributes
        .iter()
        .find(|attribute| attribute.local_name() == "spans")
    else {
        return false;
    };
    let Some((first, last)) = spans.value.trim().split_once(':') else {
        return false;
    };
    match (first.parse::<u32>(), last.parse::<u32>()) {
        (Ok(first), Ok(last)) => min < first || max > last,
        _ => false,
    }
}

fn start_tag(fragment: &[u8]) -> Result<(String, Vec<XmlAttribute>), ParseError> {
    let mut reader = Reader::from_reader(fragment);
    reader.config_mut().expand_empty_elements = false;
    loop {
        match reader.read_event().map_err(xml_err)? {
            Event::Start(element) | Event::Empty(element) => {
                let name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
                return Ok((name, attributes(&element)?));
            }
            Event::Eof => {
                return Err(ParseError::Malformed(
                    "xml fragment has no element".to_owned(),
                ));
            }
            _ => {}
        }
    }
}

fn write_start_tag(
    out: &mut Vec<u8>,
    name: &str,
    attributes: &[XmlAttribute],
    empty: bool,
) -> Result<(), ParseError> {
    let mut element = BytesStart::new(name);
    for attribute in attributes {
        element.push_attribute((attribute.name.as_str(), attribute.value.as_str()));
    }
    let mut writer = Writer::new(std::mem::take(out));
    let event = if empty {
        Event::Empty(element)
    } else {
        Event::Start(element)
    };
    writer.write_event(event).map_err(xml_err)?;
    *out = writer.into_inner();
    Ok(())
}

fn write_end_tag(out: &mut Vec<u8>, name: &str) -> Result<(), ParseError> {
    let mut writer = Writer::new(std::mem::take(out));
    writer
        .write_event(Event::End(BytesEnd::new(name)))
        .map_err(xml_err)?;
    *out = writer.into_inner();
    Ok(())
}

/// Expands a self-closing root so children can be written into it.
fn expanded_empty(
    data: &[u8],
    element: &BytesStart<'_>,
    span: Range<usize>,
) -> Result<SourceElement, ParseError> {
    let name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
    let mut prefix = Vec::new();
    write_start_tag(&mut prefix, &name, &attributes(element)?, false)?;
    let mut suffix = Vec::new();
    write_end_tag(&mut suffix, &name)?;
    suffix.extend_from_slice(&data[span.end..]);
    Ok(SourceElement {
        prefix,
        suffix,
        content: span.end..span.end,
    })
}

/// Rows and cells of a `<sheetData>`; `None` when unpatchable cell by cell.
fn scan_sheet_data(data: &[u8]) -> Result<Option<(SourceElement, Vec<SourceRow>)>, ParseError> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().expand_empty_elements = false;
    reader.config_mut().check_end_names = true;
    let mut depth = 0_usize;
    let mut prefix = Vec::new();
    let mut content_start = 0;
    let mut cursor = 0;
    let mut rows: Vec<SourceRow> = Vec::new();
    let mut row: Option<SourceRow> = None;
    let mut cell: Option<SourceCell> = None;
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(xml_err)?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                match depth {
                    0 => {
                        prefix = data[before..after].to_vec();
                        content_start = after;
                        cursor = after;
                    }
                    1 if element.local_name().as_ref() == b"row" => {
                        let Some(index) = row_index(&element, rows.last())? else {
                            return Ok(None);
                        };
                        row = Some(SourceRow {
                            before: cursor..before,
                            span: before..after,
                            tag: before..after,
                            empty: false,
                            index,
                            cells: Vec::new(),
                        });
                    }
                    2 if element.local_name().as_ref() == b"c" => {
                        if let Some(row) = &row {
                            let Some(at) = cell_address(&element, row)? else {
                                return Ok(None);
                            };
                            cell = Some(SourceCell {
                                before: row.content_start()..before,
                                span: before..after,
                                tag: before..after,
                                empty: false,
                                at,
                                shared_string: attr(&element, b"t")?.as_deref() == Some("s"),
                                formula: None,
                            });
                        }
                    }
                    3 if element.local_name().as_ref() == b"f" => {
                        if let Some(cell) = &mut cell {
                            cell.formula = Some(formula_markup(&element, before..after, false)?);
                        }
                    }
                    _ => {}
                }
                depth += 1;
                if depth > MAX_DEPTH {
                    return Err(ParseError::DepthExceeded);
                }
            }
            Event::Empty(element) => match depth {
                0 => {
                    return Ok(Some((
                        expanded_empty(data, &element, before..after)?,
                        Vec::new(),
                    )));
                }
                1 if element.local_name().as_ref() == b"row" => {
                    let Some(index) = row_index(&element, rows.last())? else {
                        return Ok(None);
                    };
                    rows.push(SourceRow {
                        before: cursor..before,
                        span: before..after,
                        tag: before..after,
                        empty: true,
                        index,
                        cells: Vec::new(),
                    });
                    cursor = after;
                }
                2 if element.local_name().as_ref() == b"c" => {
                    if let Some(row) = &mut row {
                        let Some(at) = cell_address(&element, row)? else {
                            return Ok(None);
                        };
                        let shared_string = attr(&element, b"t")?.as_deref() == Some("s");
                        row.cells.push(SourceCell {
                            before: row.content_start()..before,
                            span: before..after,
                            tag: before..after,
                            empty: true,
                            at,
                            shared_string,
                            formula: None,
                        });
                    }
                }
                3 if element.local_name().as_ref() == b"f" => {
                    if let Some(cell) = &mut cell {
                        cell.formula = Some(formula_markup(&element, before..after, true)?);
                    }
                }
                _ => {}
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                match depth {
                    0 => {
                        return Ok(Some((
                            SourceElement {
                                prefix,
                                suffix: data[before..after].to_vec(),
                                content: content_start..before,
                            },
                            rows,
                        )));
                    }
                    1 => {
                        if let Some(mut row) = row.take() {
                            row.span.end = after;
                            rows.push(row);
                            cursor = after;
                        }
                    }
                    2 => {
                        if let (Some(row), Some(mut cell)) = (&mut row, cell.take()) {
                            cell.span.end = after;
                            row.cells.push(cell);
                        }
                    }
                    _ => {}
                }
            }
            Event::Eof => {
                return Err(ParseError::Malformed("sheetData is not closed".to_owned()));
            }
            _ => {}
        }
    }
}

fn row_index(
    element: &BytesStart<'_>,
    previous: Option<&SourceRow>,
) -> Result<Option<u32>, ParseError> {
    let index = attr(element, b"r")?
        .and_then(|value| value.trim().parse::<u32>().ok())
        .filter(|value| (1..=MAX_ROWS).contains(value))
        .map(|value| value - 1);
    Ok(index.filter(|index| previous.is_none_or(|previous| previous.index < *index)))
}

fn cell_address(element: &BytesStart<'_>, row: &SourceRow) -> Result<Option<CellRef>, ParseError> {
    let at = attr(element, b"r")?.and_then(|value| CellRef::parse_a1(&value).ok());
    Ok(at.filter(|at| {
        at.row == row.index
            && row
                .cells
                .last()
                .is_none_or(|previous| previous.at.col < at.col)
    }))
}

fn formula_markup(
    element: &BytesStart<'_>,
    tag: Range<usize>,
    empty: bool,
) -> Result<FormulaMarkup, ParseError> {
    let shared = attr(element, b"t")?.as_deref() == Some("shared");
    let group = shared
        .then(|| attr(element, b"si"))
        .transpose()?
        .flatten()
        .and_then(|value| value.trim().parse::<u32>().ok());
    let reference = attr(element, b"ref")?;
    let parsed = reference
        .as_deref()
        .map(|value| CellRange::parse_a1(value).ok());
    Ok(FormulaMarkup {
        tag,
        empty,
        group,
        reference: parsed.flatten(),
        positional: attr(element, b"r1")?.is_some()
            || attr(element, b"r2")?.is_some()
            || (shared && group.is_none())
            || parsed == Some(None),
    })
}

/// Every `<col>` of a `<cols>`; `None` when one lacks a readable span.
fn scan_cols(data: &[u8]) -> Result<Option<(SourceElement, Vec<SourceCol>)>, ParseError> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().expand_empty_elements = false;
    reader.config_mut().check_end_names = true;
    let mut depth = 0_usize;
    let mut prefix = Vec::new();
    let mut content_start = 0;
    let mut cursor = 0;
    let mut cols: Vec<SourceCol> = Vec::new();
    let mut col: Option<SourceCol> = None;
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(xml_err)?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                match depth {
                    0 => {
                        prefix = data[before..after].to_vec();
                        content_start = after;
                        cursor = after;
                    }
                    1 if element.local_name().as_ref() == b"col" => {
                        let Some(range) = col_range(&element)? else {
                            return Ok(None);
                        };
                        col = Some(SourceCol {
                            before: cursor..before,
                            span: before..after,
                            tag: before..after,
                            empty: false,
                            range,
                        });
                    }
                    _ => {}
                }
                depth += 1;
                if depth > MAX_DEPTH {
                    return Err(ParseError::DepthExceeded);
                }
            }
            Event::Empty(element) => match depth {
                0 => {
                    return Ok(Some((
                        expanded_empty(data, &element, before..after)?,
                        Vec::new(),
                    )));
                }
                1 if element.local_name().as_ref() == b"col" => {
                    let Some(range) = col_range(&element)? else {
                        return Ok(None);
                    };
                    cols.push(SourceCol {
                        before: cursor..before,
                        span: before..after,
                        tag: before..after,
                        empty: true,
                        range,
                    });
                    cursor = after;
                }
                _ => {}
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                match depth {
                    0 => {
                        return Ok(Some((
                            SourceElement {
                                prefix,
                                suffix: data[before..after].to_vec(),
                                content: content_start..before,
                            },
                            cols,
                        )));
                    }
                    1 => {
                        if let Some(mut col) = col.take() {
                            col.span.end = after;
                            cols.push(col);
                            cursor = after;
                        }
                    }
                    _ => {}
                }
            }
            Event::Eof => {
                return Err(ParseError::Malformed("cols is not closed".to_owned()));
            }
            _ => {}
        }
    }
}

/// The zero-based columns a `<col>` covers, clamped as the parser clamps them.
fn col_range(element: &BytesStart<'_>) -> Result<Option<Range<u32>>, ParseError> {
    let min = attr(element, b"min")?.and_then(|value| value.trim().parse::<u32>().ok());
    let max = attr(element, b"max")?.and_then(|value| value.trim().parse::<u32>().ok());
    let (Some(min), Some(max)) = (min, max) else {
        return Ok(None);
    };
    let min = min.clamp(1, MAX_COLS);
    let max = max.clamp(min, MAX_COLS);
    Ok(Some(min - 1..max))
}
