//! Lexical patching of a preserved worksheet's `<cols>` and `<sheetData>`.

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::iter::Peekable;
use std::ops::{Range, RangeInclusive};

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::name::ResolveResult;
use quick_xml::{NsReader, Reader, Writer};
use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::{Cell, CellRange, CellRef, CellValue, ColId, ColStyle, Sheet, Workbook};

use crate::axis::SheetAxes;
use crate::future_functions::FormulaNames;
use crate::package::{XmlAttribute, attributes, remove_attribute, set_attribute};
use crate::read::{SharedStringCells, SourceCellFacts, WorksheetCellFacts, is_truthy};
use crate::write::{
    SharedStringPlan, cell_value_markup, col_runs, fmt_num, fragment, shared_string_index,
    write_cell, write_col_run, write_cols, write_row,
};
use crate::xml::{attr, resolve_entity, text_attribute, xml_err};
use crate::{MAX_DEPTH, ParseError};

#[cfg(test)]
mod cache_patch_tests;
#[cfg(any(test, feature = "test-oracle"))]
#[cfg_attr(not(test), allow(dead_code))]
mod oracle;
#[cfg(test)]
mod oracle_tests;
#[cfg(test)]
mod regeneration_tests;
mod style_match;
#[cfg(test)]
mod style_match_tests;

pub(crate) use style_match::{SourceStyles, StyleMatch, keeps_source_xf, retains_source_styles};
pub use style_match::{keeps_source, keeps_source_conditional_formats};

type SharedGroupEntry = (bool, Option<CellRange>, Option<(u32, u32)>);

#[cfg(any(test, feature = "test-oracle"))]
thread_local! {
    static EMISSION_LOOKUPS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    static DETECTOR_LOOKUPS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

/// One edited sheet against the source it was read from.
pub(crate) struct SheetPatch<'a> {
    pub(crate) sheet: &'a Sheet,
    pub(crate) original: &'a Sheet,
    pub(crate) cell_facts: &'a WorksheetCellFacts,
    pub(crate) axes: &'a SheetAxes,
    pub(crate) workbook: &'a Workbook,
    pub(crate) sst_index: &'a HashMap<&'a str, usize>,
    pub(crate) retained: &'a SharedStringCells,
    pub(crate) plan: Option<&'a SharedStringPlan>,
    pub(crate) styles: &'a StyleMatch<'a>,
    /// The defined names this sheet's formulas resolve.
    pub(crate) names: FormulaNames<'a>,
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
    alternates: Vec<SourceAlternate>,
}

impl SourceRow {
    fn content_start(&self) -> usize {
        let cell_end = self.cells.last().map_or(self.tag.end, |cell| cell.span.end);
        self.alternates
            .last()
            .map_or(cell_end, |block| cell_end.max(block.end))
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
    changed: Option<(Option<f64>, Option<f64>)>,
    style: Option<Option<u32>>,
}

struct SourceSheetData {
    element: SourceElement,
    rows: Vec<SourceRow>,
    alternates: Vec<SourceAlternate>,
}

struct SourceAlternate {
    before: usize,
    end: usize,
    indices: RangeInclusive<u32>,
}

impl SourceAlternate {
    fn emit(&self, out: &mut Vec<u8>, data: &[u8]) {
        out.extend_from_slice(&data[self.before..self.end]);
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct AlternateSpan {
    pub(crate) rows: Range<u32>,
    pub(crate) cols: Range<u32>,
    pub(crate) wraps_rows: bool,
}

#[derive(Default, PartialEq, Eq)]
struct AlternateBranch {
    cells: BTreeSet<(u32, u32)>,
    rows: BTreeSet<u32>,
}

struct AlternateBlock {
    depth: usize,
    parent_row: Option<u32>,
    branch: Option<AlternateBranch>,
    first_branch: Option<AlternateBranch>,
    all_cells: BTreeSet<(u32, u32)>,
    all_rows: BTreeSet<u32>,
    active_row: Option<u32>,
    inert: bool,
    has_choice: bool,
    has_fallback: bool,
}

impl AlternateBlock {
    fn finish_branch(&mut self) {
        let branch = self.branch.take().expect("opened branch");
        self.all_cells.extend(branch.cells.iter().copied());
        self.all_rows.extend(branch.rows.iter().copied());
        if let Some(first) = &self.first_branch {
            self.inert &= *first == branch;
        } else {
            self.first_branch = Some(branch);
        }
    }

    fn finish(self, facts: &mut SourceCellFacts) {
        let span = |min: Option<u32>, max: Option<u32>| {
            min.zip(max).map_or(0..0, |(min, max)| min..max + 1)
        };
        let rows = self.parent_row.map_or_else(
            || {
                span(
                    self.all_rows.first().copied(),
                    self.all_rows.last().copied(),
                )
            },
            |row| row..row + 1,
        );
        let cols = span(
            self.all_cells.iter().map(|&(_, col)| col).min(),
            self.all_cells.iter().map(|&(_, col)| col).max(),
        );
        facts.non_inert_alternate |= !self.inert
            || !self.has_choice
            || rows.is_empty()
            || (self.parent_row.is_some() && cols.is_empty());
        facts.alternate_cells.extend(self.all_cells);
        facts.alternate_rows.extend(self.all_rows);
        facts.alternate_spans.push(AlternateSpan {
            rows,
            cols,
            wraps_rows: self.parent_row.is_none(),
        });
    }
}

pub(crate) fn scan_alternate_content(
    data: &[u8],
    namespace_stable: bool,
    facts: &mut SourceCellFacts,
) -> Result<(), ParseError> {
    let mut reader = NsReader::from_reader(data);
    let mut depth = 0;
    let mut worksheet_ns = None;
    let mut grid = false;
    let mut row = None;
    let mut row_namespace_stable = true;
    let mut block: Option<AlternateBlock> = None;
    let mut formula: Option<(usize, String)> = None;
    loop {
        let (ns, event) = reader.read_resolved_event().map_err(xml_err)?;
        match event {
            Event::Start(ref tag) | Event::Empty(ref tag) => {
                let empty = matches!(event, Event::Empty(_));
                let name = tag.local_name();
                let name = name.as_ref();
                let namespace = match ns {
                    ResolveResult::Bound(ns) => Some(ns.as_ref().to_vec()),
                    ResolveResult::Unbound => None,
                    ResolveResult::Unknown(_) => Some(Vec::new()),
                };
                let mc = namespace.as_deref()
                    == Some(crate::drawing_anchors::NS_MARKUP_COMPATIBILITY.as_bytes());
                if let Some(block) = &mut block {
                    if namespace.as_deref().is_some_and(<[u8]>::is_empty) {
                        block.inert = false;
                    }
                    if name == b"f" {
                        facts.sheet_data_alternate_formulas = true;
                        block.inert = false;
                        if !empty && formula.is_none() {
                            formula = Some((depth, String::new()));
                        }
                    } else if name == b"AlternateContent" {
                        block.inert = false;
                    }
                    if depth == block.depth + 1 {
                        let choice = mc && name == b"Choice" && !block.has_fallback;
                        let fallback = mc && name == b"Fallback" && !block.has_fallback;
                        block.inert &= choice || fallback;
                        block.has_choice |= choice;
                        block.has_fallback |= fallback;
                        block.branch = Some(AlternateBranch::default());
                        block.active_row = None;
                        if empty {
                            block.finish_branch();
                        }
                    } else if let Some(branch) = &mut block.branch {
                        let child = depth == block.depth + 2;
                        let wraps_rows = block.parent_row.is_none();
                        if child && wraps_rows && name == b"row" && namespace == worksheet_ns {
                            let index = row_index(tag, None)?;
                            block.inert &= index.is_some_and(|index| {
                                let ordered = branch.rows.last().is_none_or(|last| *last < index);
                                branch.rows.insert(index);
                                ordered
                            });
                            block.active_row = if empty { None } else { index };
                        } else if name == b"c"
                            && namespace == worksheet_ns
                            && (child && !wraps_rows || depth == block.depth + 3 && wraps_rows)
                        {
                            let at = attr(tag, b"r")?.and_then(|r| CellRef::parse_a1(&r).ok());
                            block.inert &= at.is_some_and(|at| {
                                let key = (at.row, at.col);
                                let ordered = branch.cells.last().is_none_or(|last| *last < key);
                                branch.cells.insert(key);
                                ordered && Some(at.row) == block.parent_row.or(block.active_row)
                            });
                        } else if child || name == b"row" || name == b"c" {
                            block.inert = false;
                        }
                    }
                } else if depth == 0 {
                    worksheet_ns = namespace;
                } else if depth == 1 && name == b"sheetData" && namespace == worksheet_ns {
                    grid = !empty;
                    row = None;
                } else if grid && name == b"AlternateContent" {
                    if empty {
                        facts.non_inert_alternate = true;
                    } else {
                        block = Some(AlternateBlock {
                            depth,
                            parent_row: row,
                            branch: None,
                            first_branch: None,
                            all_cells: BTreeSet::new(),
                            all_rows: BTreeSet::new(),
                            active_row: None,
                            inert: mc
                                && namespace_stable
                                && (depth == 2
                                    || (depth == 3 && row.is_some() && row_namespace_stable)),
                            has_choice: false,
                            has_fallback: false,
                        });
                    }
                } else if grid && depth == 2 && name == b"row" && !empty {
                    row = row_index(tag, None)?;
                    row_namespace_stable = attr(tag, b"xmlns")?.is_none_or(|value| {
                        value.as_bytes() == worksheet_ns.as_deref().unwrap_or_default()
                    });
                }
                depth += usize::from(!empty);
            }
            Event::Text(text) => {
                if let Some((_, value)) = &mut formula {
                    value.push_str(&text.xml10_content().map_err(xml_err)?);
                }
            }
            Event::CData(text) => {
                if let Some((_, value)) = &mut formula {
                    value.push_str(&text.xml10_content().map_err(xml_err)?);
                }
            }
            Event::GeneralRef(reference) => {
                if let Some((_, value)) = &mut formula {
                    value.push_str(&resolve_entity(&reference.decode().map_err(xml_err)?)?);
                }
            }
            Event::End(_) => {
                depth -= 1;
                if let Some((_, text)) = formula.take_if(|(opened, _)| *opened == depth) {
                    facts.alternate_formula_texts.push(text);
                }
                if let Some(open) = &mut block {
                    if depth == open.depth + 1 && open.branch.is_some() {
                        open.finish_branch();
                    }
                    if depth == open.depth {
                        block.take().expect("opened wrapper").finish(facts);
                    }
                } else if depth == 1 {
                    grid = false;
                } else if depth == 2 {
                    row = None;
                }
            }
            Event::Eof => return Ok(()),
            _ => {}
        }
    }
}

struct SourceCellCursor<'a, I: Iterator> {
    sheet: &'a Sheet,
    cells: Peekable<I>,
    source: Option<(u32, u32)>,
    current: Option<(u32, u32)>,
}

impl<'a, I> SourceCellCursor<'a, I>
where
    I: Iterator<Item = (CellRef, &'a Cell)>,
{
    fn new(sheet: &'a Sheet, cells: I) -> Self {
        Self {
            sheet,
            cells: cells.peekable(),
            source: None,
            current: None,
        }
    }

    fn at(&mut self, source: CellRef, current: CellRef) -> (Option<&'a Cell>, bool) {
        let source_key = (source.row, source.col);
        let current_key = (current.row, current.col);
        let source_ordered = self.source.is_none_or(|previous| previous < source_key);
        let current_ordered = self.current.is_none_or(|previous| previous < current_key);
        self.source = Some(self.source.map_or(source_key, |last| last.max(source_key)));
        self.current = Some(
            self.current
                .map_or(current_key, |last| last.max(current_key)),
        );
        let original = if source_ordered {
            while self
                .cells
                .peek()
                .is_some_and(|(at, _)| (at.row, at.col) < source_key)
            {
                self.cells.next();
            }
            self.cells
                .next_if(|(at, _)| (at.row, at.col) == source_key)
                .map(|(_, cell)| cell)
        } else {
            #[cfg(test)]
            EMISSION_LOOKUPS.with(|count| count.set(count.get() + 1));
            self.sheet.cell(source)
        };
        (original, source_ordered && current_ordered)
    }
}

impl SheetPatch<'_> {
    /// Patched `<sheetData>`; `None` when the source resists cell-by-cell patching.
    pub(crate) fn sheet_data(&self, source: &[u8]) -> Result<Option<Vec<u8>>, ParseError> {
        let Some(SourceSheetData {
            element,
            rows,
            alternates,
        }) = scan_sheet_data(source)?
        else {
            return Ok(None);
        };
        #[cfg(any(test, feature = "test-oracle"))]
        if crate::save_oracle::use_legacy_save_path()
            && alternates.is_empty()
            && rows.iter().all(|row| row.alternates.is_empty())
        {
            return self.sheet_data_oracle(source);
        }
        let wrapped = &self.cell_facts.source;
        let dirty = self.dirty_formulas(source, &rows)?;
        let mut out = Vec::with_capacity(source.len());
        out.extend_from_slice(&element.prefix);
        let mut source_rows: Vec<_> = rows
            .iter()
            .filter_map(|row| {
                self.axes
                    .rows
                    .current(row.index)
                    .map(|current| (current, row))
            })
            .collect();
        // a move makes the map non-monotone; the element is written in current
        // order, so the rows it reuses are visited in that order too
        source_rows.sort_unstable_by_key(|(current, _)| *current);
        let mut source_rows = source_rows.into_iter().peekable();
        let mut alternate_blocks = alternates.iter().peekable();
        let mut cells = self
            .sheet
            .iter_cells()
            .filter(|(at, _)| !wrapped.alternate_cells.contains(&(at.row, at.col)))
            .peekable();
        let mut originals = SourceCellCursor::new(self.original, self.original.iter_cells());
        let mut heights = self
            .sheet
            .row_heights
            .keys()
            .copied()
            .filter(|row| !wrapped.alternate_rows.contains(row))
            .peekable();
        let mut row_styles = self
            .sheet
            .row_styles
            .keys()
            .copied()
            .filter(|row| !wrapped.alternate_rows.contains(row))
            .peekable();
        loop {
            let next_source = source_rows.peek().map(|(current, _)| *current);
            let next_alternate = alternate_blocks.peek().map(|block| *block.indices.start());
            let next_cell = cells.peek().map(|(at, _)| at.row);
            let next_height = heights.peek().copied();
            let next_style = row_styles.peek().copied();
            let Some(row) = [
                next_source,
                next_cell,
                next_height,
                next_style,
                next_alternate,
            ]
            .into_iter()
            .flatten()
            .min() else {
                break;
            };
            if next_alternate == Some(row) {
                alternate_blocks
                    .next()
                    .expect("peeked")
                    .emit(&mut out, source);
                continue;
            }
            if next_height == Some(row) {
                heights.next();
            }
            if next_style == Some(row) {
                row_styles.next();
            }
            if next_source == Some(row) {
                let (_, source_row) = source_rows.next().expect("peeked");
                self.emit_source_row(
                    &mut out,
                    source,
                    source_row,
                    row,
                    &mut cells,
                    &mut originals,
                    &dirty,
                )?;
            } else {
                self.emit_generated_row(&mut out, row, &mut cells)?;
            }
            let current = (row, u32::MAX);
            originals.current = Some(originals.current.map_or(current, |last| last.max(current)));
        }
        let tail = rows
            .last()
            .map_or(element.content.start, |row| row.span.end);
        let tail = alternates.last().map_or(tail, |block| tail.max(block.end));
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
        let mut entries: Vec<(u32, Vec<u8>)> = Vec::new();
        let mut covered = Vec::new();
        let model_runs = self.sheet.col_style_runs();
        let source_runs = self.original.col_style_runs();
        for col in &cols {
            let (name, attributes) = start_tag(&source[col.tag.clone()])?;
            let pieces = self
                .axes
                .cols
                .current_ranges(col.range.clone())
                .into_iter()
                .flat_map(|range| self.width_pieces(range, &model_runs, &source_runs))
                .collect::<Vec<_>>();
            let verbatim = pieces.len() == 1
                && pieces[0].range == col.range
                && pieces[0].changed.is_none()
                && pieces[0].style.is_none();
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
                    if let Some((width, hidden_width)) = piece.changed {
                        set_col_width(&mut attributes, width, hidden_width);
                    }
                    match piece.style {
                        Some(Some(style)) => {
                            set_attribute(&mut attributes, "style", "style", style.to_string())
                        }
                        Some(None) => remove_attribute(&mut attributes, "style"),
                        None => {}
                    }
                    write_start_tag(&mut bytes, &name, &attributes, col.empty)?;
                    if !col.empty {
                        bytes.extend_from_slice(&source[col.tag.end..col.span.end]);
                    }
                }
                covered.push(piece.range.clone());
                entries.push((piece.range.start, bytes));
            }
        }
        let mut generated = col_runs(self.sheet, &covered).into_iter().peekable();
        if entries.is_empty() && generated.peek().is_none() {
            return Ok(None);
        }
        let mut out = element.prefix;
        for (start, bytes) in entries {
            while let Some(run) = generated.next_if(|run| run.first < start) {
                write_col_run(&mut Writer::new(&mut out), &run).map_err(xml_err)?;
            }
            out.extend_from_slice(&bytes);
        }
        for run in generated {
            write_col_run(&mut Writer::new(&mut out), &run).map_err(xml_err)?;
        }
        let tail = cols
            .last()
            .map_or(element.content.start, |col| col.span.end);
        out.extend_from_slice(&source[tail..element.content.end]);
        out.extend_from_slice(&element.suffix);
        Ok(Some(out))
    }

    fn generated_cols(&self) -> Result<Option<Vec<u8>>, ParseError> {
        (!self.sheet.col_widths.is_empty() || !self.sheet.col_styles.is_empty())
            .then(|| fragment(|writer| write_cols(writer, self.sheet)))
            .transpose()
    }

    /// Splits a column span wherever the model width or style parted from the
    /// source; the runs are both sides' ascending, disjoint column styles.
    fn width_pieces(
        &self,
        range: Range<u32>,
        model_runs: &[ColStyle],
        source_runs: &[ColStyle],
    ) -> Vec<WidthPiece> {
        let mut pieces: Vec<WidthPiece> = Vec::new();
        for col in range {
            let source = self.axes.cols.source(col);
            let model = self.sheet.col_widths.get(&col).copied();
            let model_hidden = self.sheet.hidden_col_widths.get(&col).copied();
            let original = source.and_then(|source| self.original.col_widths.get(&source).copied());
            let original_hidden =
                source.and_then(|source| self.original.hidden_col_widths.get(&source).copied());
            let changed = ((model, model_hidden) != (original, original_hidden))
                .then_some((model, model_hidden));
            let model_style = run_style(model_runs, col);
            let source_style = source.and_then(|source| run_style(source_runs, source));
            let same = self.styles.equivalent(source_style, model_style);
            let style = (!same).then_some(model_style);
            match pieces.last_mut() {
                Some(piece) if piece.changed == changed && piece.style == style => {
                    piece.range.end = col + 1
                }
                _ => pieces.push(WidthPiece {
                    range: col..col + 1,
                    changed,
                    style,
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
            None,
            self.names,
        )
        .map_err(xml_err)?;
        *out = writer.into_inner();
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn emit_source_row<'c, 'o, I, O>(
        &self,
        out: &mut Vec<u8>,
        data: &[u8],
        source: &SourceRow,
        row: u32,
        cells: &mut Peekable<I>,
        originals: &mut SourceCellCursor<'o, O>,
        dirty: &DirtyFormulas,
    ) -> Result<(), ParseError>
    where
        I: Iterator<Item = (CellRef, &'c Cell)>,
        O: Iterator<Item = (CellRef, &'o Cell)>,
    {
        let mut body = Vec::new();
        let mut alternates = source.alternates.iter().peekable();
        let mut columns: Option<(u32, u32)> = None;
        let mapped = source.cells.iter().filter_map(|cell| {
            self.axes
                .cols
                .current(cell.at.col)
                .map(|current| (current, cell))
        });
        let mut sorted = Vec::new();
        let mut streaming = None;
        if source.alternates.is_empty() {
            streaming = Some(mapped);
        } else {
            sorted.extend(mapped);
            sorted.sort_unstable_by_key(|(current, _)| *current);
        }
        let mut source_cells = streaming.into_iter().flatten().chain(sorted).peekable();
        loop {
            let next_source = source_cells.peek().map(|(current, _)| *current);
            let next_alternate = alternates.peek().map(|block| *block.indices.start());
            let next_model = cells
                .peek()
                .filter(|(at, _)| at.row == row)
                .map(|(at, _)| at.col);
            let Some(col) = [next_source, next_model, next_alternate]
                .into_iter()
                .flatten()
                .min()
            else {
                break;
            };
            if next_alternate == Some(col) {
                alternates.next().expect("peeked").emit(&mut body, data);
                continue;
            }
            let source_cell =
                (next_source == Some(col)).then(|| source_cells.next().expect("peeked").1);
            let model_cell = (next_model == Some(col)).then(|| cells.next().expect("peeked").1);
            let at = CellRef::new(row, col);
            let mut original = None;
            let verbatim = source_cell.is_some_and(|source_cell| {
                let (found, ordered) = originals.at(source_cell.at, at);
                original = found;
                let model = if ordered {
                    model_cell
                } else {
                    #[cfg(test)]
                    EMISSION_LOOKUPS.with(|count| count.set(count.get() + 1));
                    self.sheet.cell(at)
                };
                self.verbatim(source_cell, at, model, found, dirty)
            });
            match source_cell {
                Some(source_cell) if verbatim => {
                    body.extend_from_slice(&data[source_cell.before.clone()]);
                    self.emit_source_cell(&mut body, data, source_cell, at)?;
                }
                Some(source_cell) => {
                    let Some(cell) = model_cell else {
                        continue;
                    };
                    body.extend_from_slice(&data[source_cell.before.clone()]);
                    if !self.emit_cached_cell(
                        &mut body,
                        data,
                        source_cell,
                        at,
                        cell,
                        original,
                        dirty,
                    )? {
                        self.emit_cell(&mut body, at, cell, original)?;
                    }
                }
                None => self.emit_cell(
                    &mut body,
                    at,
                    model_cell.expect("one side is present"),
                    None,
                )?,
            }
            columns = Some(match columns {
                Some((min, max)) => (min, max.max(col + 1)),
                None => (col + 1, col + 1),
            });
        }

        let moved = source.index != row;
        let height = self.sheet.row_heights.get(&row).copied();
        let hidden_height = self.sheet.hidden_row_heights.get(&row).copied();
        let height_changed = (height, hidden_height)
            != (
                self.original.row_heights.get(&source.index).copied(),
                self.original.hidden_row_heights.get(&source.index).copied(),
            );
        let style = self.sheet.row_styles.get(&row).copied();
        let style_changed = style != self.original.row_styles.get(&source.index).copied();
        let (name, mut attributes) = start_tag(&data[source.tag.clone()])?;
        let mut rewrite = moved || height_changed || style_changed;
        if moved {
            set_attribute(&mut attributes, "r", "r", (u64::from(row) + 1).to_string());
        }
        if height_changed {
            set_row_height(&mut attributes, height, hidden_height);
        }
        if style_changed {
            remove_attribute(&mut attributes, "s");
            remove_attribute(&mut attributes, "customFormat");
            if let Some(style) = style {
                set_attribute(&mut attributes, "s", "s", style.to_string());
                set_attribute(&mut attributes, "customFormat", "customFormat", "1".into());
            }
        }
        let mut wrapped_cols = self
            .cell_facts
            .source
            .alternate_cells
            .range((row, 0)..=(row, u32::MAX));
        if let Some(&(_, first)) = wrapped_cols.next() {
            let last = wrapped_cols.next_back().map_or(first, |&(_, last)| last);
            let (min, max) = columns.unwrap_or((first + 1, last + 1));
            columns = Some((min.min(first + 1), max.max(last + 1)));
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
    fn verbatim(
        &self,
        cell: &SourceCell,
        at: CellRef,
        model: Option<&Cell>,
        original: Option<&Cell>,
        dirty: &DirtyFormulas,
    ) -> bool {
        if !self.styles.same(original, model) {
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
        self.formula_clean(cell, at, formula, dirty)
    }

    fn formula_clean(
        &self,
        cell: &SourceCell,
        at: CellRef,
        formula: &FormulaMarkup,
        dirty: &DirtyFormulas,
    ) -> bool {
        if cell.at != at && formula.positional {
            return false;
        }
        match formula.group {
            Some(group) => !dirty.groups.contains(&group),
            None => !dirty.masters.contains(&(cell.at.row, cell.at.col)),
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
        self.emit_source_cell_content(out, data, cell, at, cell.span.end)
    }

    fn emit_source_cell_content(
        &self,
        out: &mut Vec<u8>,
        data: &[u8],
        cell: &SourceCell,
        at: CellRef,
        end: usize,
    ) -> Result<(), ParseError> {
        let mut cursor = cell.tag.end;
        if cell.at != at
            && let Some(formula) = &cell.formula
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
        out.extend_from_slice(&data[cursor..end]);
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn emit_cached_cell(
        &self,
        out: &mut Vec<u8>,
        data: &[u8],
        source: &SourceCell,
        at: CellRef,
        cell: &Cell,
        original: Option<&Cell>,
        dirty: &DirtyFormulas,
    ) -> Result<bool, ParseError> {
        let (Some(formula), Some(original)) = (&source.formula, original) else {
            return Ok(false);
        };
        if cell.formula.is_none()
            || cell.formula != original.formula
            || !self.formula_clean(source, at, formula, dirty)
        {
            return Ok(false);
        }
        let Some(slot) = self.cache_patch_slot(data, source, at)? else {
            return Ok(false);
        };
        let (name, mut attributes) = start_tag(&data[source.tag.clone()])?;
        if source.at != at {
            set_attribute(&mut attributes, "r", "r", at.to_a1());
        }
        let style = self.styles.written(original.style, cell.style);
        if style != original.style {
            attributes.retain(|attribute| attribute.name != "s");
            if let Some(style) = style {
                attributes.push(XmlAttribute {
                    name: "s".to_owned(),
                    value: style.to_string(),
                });
            }
        }
        let markup = cell_value_markup(cell, self.sst_index, None);
        match markup.ty {
            Some(ty) => {
                if let Some(attribute) = attributes
                    .iter_mut()
                    .find(|attribute| attribute.name == "t")
                {
                    attribute.value = ty.to_owned();
                } else {
                    attributes.push(XmlAttribute {
                        name: "t".to_owned(),
                        value: ty.to_owned(),
                    });
                }
            }
            None => attributes.retain(|attribute| attribute.name != "t"),
        }
        write_start_tag(out, &name, &attributes, false)?;
        self.emit_source_cell_content(out, data, source, at, slot.start)?;
        if let Some(value) = markup.value {
            let mut writer = Writer::new(std::mem::take(out));
            writer
                .create_element("v")
                .write_text_content(BytesText::new(&value))
                .map_err(xml_err)?;
            *out = writer.into_inner();
        }
        out.extend_from_slice(&data[slot.end..source.span.end]);
        Ok(true)
    }

    fn cache_patch_slot(
        &self,
        data: &[u8],
        source: &SourceCell,
        at: CellRef,
    ) -> Result<Option<Range<usize>>, ParseError> {
        let Some(formula) = &source.formula else {
            return Ok(None);
        };
        let (name, attributes) = start_tag(&data[source.tag.clone()])?;
        if name != "c"
            || attributes
                .iter()
                .any(|attribute| attribute.name.contains(':') || attribute.name == "xmlns")
        {
            return Ok(None);
        }
        let (_, formula_attributes) = start_tag(&data[formula.tag.clone()])?;
        let metadata = attributes
            .iter()
            .any(|attribute| matches!(attribute.local_name(), "cm" | "vm"));
        let array = formula_attributes
            .iter()
            .any(|attribute| attribute.local_name() == "t" && attribute.value == "array");
        if metadata || array {
            let source_range = self.original.array_formula(source.at);
            if source_range.is_none()
                || source_range != formula.reference
                || source_range != self.sheet.array_formula(at)
            {
                return Ok(None);
            }
        }
        cached_value_slot(data, source)
    }

    fn emit_cell(
        &self,
        out: &mut Vec<u8>,
        at: CellRef,
        cell: &Cell,
        original: Option<&Cell>,
    ) -> Result<(), ParseError> {
        let retained = shared_string_index(
            cell,
            at,
            self.workbook,
            self.sst_index,
            self.retained,
            self.plan,
        );
        let mut writer = Writer::new(std::mem::take(out));
        let style = if original.is_some()
            || self
                .axes
                .rows
                .source(at.row)
                .zip(self.axes.cols.source(at.col))
                .is_some()
        {
            self.styles
                .written(original.and_then(|cell| cell.style), cell.style)
        } else {
            cell.style
        };
        write_cell(
            &mut writer,
            at,
            cell,
            style,
            self.sst_index,
            retained,
            self.sheet.array_formula(at),
            self.names,
        )
        .map_err(xml_err)?;
        *out = writer.into_inner();
        Ok(())
    }

    fn dirty_formulas(&self, data: &[u8], rows: &[SourceRow]) -> Result<DirtyFormulas, ParseError> {
        let mut changed = self.changed_source_cells();
        let mut arrays = BTreeSet::new();
        let mut groups: HashMap<u32, SharedGroupEntry> = HashMap::new();
        let mut members: HashMap<u32, Vec<CellRef>> = HashMap::new();
        let mut masters = Vec::new();
        for cell in rows.iter().flat_map(|row| &row.cells) {
            let Some(formula) = &cell.formula else {
                continue;
            };
            let key = (cell.at.row, cell.at.col);
            if self
                .original
                .array_formula(cell.at)
                .and_then(|range| self.remap_range(range))
                != self
                    .mapped(cell.at)
                    .and_then(|mapped| self.sheet.array_formula(mapped))
            {
                changed.insert(key);
                arrays.insert(key);
            }
            match formula.group {
                Some(group) => {
                    let entry = groups.entry(group).or_insert((false, None, None));
                    entry.0 |= self
                        .mapped(cell.at)
                        .is_none_or(|at| self.inverse(at) != Some(cell.at))
                        || (changed.contains(&key) && self.source_formula_changed(cell.at));
                    entry.0 |= arrays.contains(&key);
                    if let Some(at) = self.mapped(cell.at) {
                        let shift = (
                            at.row.wrapping_sub(cell.at.row),
                            at.col.wrapping_sub(cell.at.col),
                        );
                        entry.0 |= *entry.2.get_or_insert(shift) != shift;
                        members.entry(group).or_default().push(at);
                    }
                    if !entry.0
                        && formula.reference.is_some()
                        && let Some(at) = self.mapped(cell.at)
                    {
                        entry.0 |= (cell.at != at && formula.positional)
                            || (changed.contains(&key)
                                && self.cache_patch_slot(data, cell, at)?.is_none());
                    }
                    if let Some(reference) = formula.reference
                        && entry.1.replace(reference).is_some()
                    {
                        entry.0 = true;
                    }
                }
                None => {
                    masters.push((key, formula.reference));
                }
            }
        }
        Ok(DirtyFormulas {
            groups: groups
                .into_iter()
                .filter(|(group, (dirty, master, _))| {
                    let members = members.get(group).map_or(&[][..], Vec::as_slice);
                    *dirty
                        || !master
                            .filter(|&reference| self.moves_uniformly(reference))
                            .and_then(|reference| self.remap_range(reference))
                            .is_some_and(|written| members.iter().all(|&at| written.contains(at)))
                })
                .map(|(group, _)| group)
                .collect(),
            masters: self.dirty_masters(masters, &arrays, &changed),
        })
    }

    fn dirty_masters(
        &self,
        masters: Vec<((u32, u32), Option<CellRange>)>,
        arrays: &BTreeSet<(u32, u32)>,
        changed: &BTreeSet<(u32, u32)>,
    ) -> HashSet<(u32, u32)> {
        let mut changed_rows = None;
        masters
            .into_iter()
            .filter(|(anchor, reference)| {
                arrays.contains(anchor)
                    || reference.is_some_and(|reference| {
                        changed.contains(anchor)
                            || !self.moves_uniformly(reference)
                            || self.ranged_master_dirty(
                                *anchor,
                                reference,
                                changed_rows.get_or_insert_with(|| {
                                    let mut rows = BTreeMap::<u32, Vec<u32>>::new();
                                    for &(row, col) in changed {
                                        rows.entry(row).or_default().push(col);
                                    }
                                    rows
                                }),
                            )
                    })
            })
            .map(|(anchor, _)| anchor)
            .collect()
    }

    fn ranged_master_dirty(
        &self,
        anchor: (u32, u32),
        reference: CellRange,
        changed_rows: &BTreeMap<u32, Vec<u32>>,
    ) -> bool {
        let mut rows = changed_rows
            .range(reference.start.row..=reference.end.row)
            .peekable();
        if rows.peek().is_none() {
            return false;
        }
        let stable_array = self.axes.is_identity()
            && self
                .original
                .array_formula(CellRef::new(anchor.0, anchor.1))
                == Some(reference);
        rows.any(|(&row, cols)| {
            let start = cols.partition_point(|&col| col < reference.start.col);
            cols[start..]
                .iter()
                .take_while(|&&col| col <= reference.end.col)
                .any(|&col| {
                    xlsx_model::open_timing! {
                        xlsx_model::open_timing::count("save_array_range_changes", 1);
                    }
                    !stable_array || !self.fills_empty_cache(CellRef::new(row, col))
                })
        })
    }

    fn fills_empty_cache(&self, at: CellRef) -> bool {
        let Some(cell) = self.sheet.cell(at) else {
            return false;
        };
        let original = self.original.cell(at);
        cell.value != CellValue::Empty
            && !self.cell_facts.has_cached_value(at, original)
            && match original {
                Some(source) => {
                    source.formula == cell.formula
                        && self.styles.equivalent(source.style, cell.style)
                }
                None => cell.formula.is_none() && cell.style.is_none(),
            }
    }

    fn source_formula_changed(&self, source: CellRef) -> bool {
        let Some(at) = self.mapped(source) else {
            return true;
        };
        if self.inverse(at) != Some(source) {
            return true;
        }
        match (self.original.cell(source), self.sheet.cell(at)) {
            (Some(original), Some(cell)) => original.formula != cell.formula,
            _ => true,
        }
    }

    /// Source addresses whose cell changed or is gone.
    fn changed_source_cells(&self) -> BTreeSet<(u32, u32)> {
        #[cfg(any(test, feature = "test-oracle"))]
        if crate::save_oracle::use_legacy_save_path() {
            return self.changed_source_cells_oracle();
        }
        let mut changed = BTreeSet::new();
        let mut cells = self.sheet.iter_cells().peekable();
        let mut previous = None;
        for (source, original) in self.original.iter_cells() {
            let Some(mapped) = self.mapped(source) else {
                changed.insert((source.row, source.col));
                continue;
            };
            let key = (mapped.row, mapped.col);
            if previous.is_some_and(|previous| previous >= key)
                || self.inverse(mapped) != Some(source)
            {
                return self.changed_source_cells_by_lookup();
            }
            previous = Some(key);
            while cells.peek().is_some_and(|(at, _)| (at.row, at.col) < key) {
                let (at, _) = cells.next().expect("peeked");
                if !self.record_unmatched_source(at, &mut changed) {
                    return self.changed_source_cells_by_lookup();
                }
            }
            let current = cells
                .next_if(|(at, _)| (at.row, at.col) == key)
                .map(|(_, cell)| cell);
            if !self.styles.same(Some(original), current) {
                changed.insert((source.row, source.col));
            }
        }
        for (at, _) in cells {
            if !self.record_unmatched_source(at, &mut changed) {
                return self.changed_source_cells_by_lookup();
            }
        }
        changed
    }

    fn record_unmatched_source(&self, at: CellRef, changed: &mut BTreeSet<(u32, u32)>) -> bool {
        if let Some(source) = self.inverse(at) {
            if self.mapped(source) != Some(at) {
                return false;
            }
            changed.insert((source.row, source.col));
        }
        true
    }

    fn changed_source_cells_by_lookup(&self) -> BTreeSet<(u32, u32)> {
        let mut changed = BTreeSet::new();
        for (at, cell) in self.original.iter_cells() {
            let current = self.mapped(at).and_then(|mapped| {
                #[cfg(test)]
                DETECTOR_LOOKUPS.with(|count| count.set(count.get() + 1));
                self.sheet.cell(mapped)
            });
            if !self.styles.same(Some(cell), current) {
                changed.insert((at.row, at.col));
            }
        }
        for (at, cell) in self.sheet.iter_cells() {
            if let Some(source) = self.inverse(at) {
                #[cfg(test)]
                DETECTOR_LOOKUPS.with(|count| count.set(count.get() + 1));
                if !self.styles.same(self.original.cell(source), Some(cell)) {
                    changed.insert((source.row, source.col));
                }
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

fn cached_value_slot(data: &[u8], cell: &SourceCell) -> Result<Option<Range<usize>>, ParseError> {
    let mut reader = Reader::from_reader(&data[cell.span.clone()]);
    reader.config_mut().expand_empty_elements = false;
    let mut depth = 0_usize;
    let mut child_start = cell.tag.end;
    let mut formula_end = None;
    let mut value = None;
    loop {
        let before = cell.span.start + reader.buffer_position() as usize;
        let event = reader.read_event().map_err(xml_err)?;
        let after = cell.span.start + reader.buffer_position() as usize;
        if let Event::Start(element) | Event::Empty(element) = &event {
            if depth > 1 {
                return Ok(None);
            }
            if depth == 1
                && (!matches!(element.name().as_ref(), b"f" | b"v" | b"is")
                    || attributes(element)?.iter().any(|attribute| {
                        attribute.name == "xmlns" || attribute.name.starts_with("xmlns:")
                    }))
            {
                return Ok(None);
            }
        }
        match event {
            Event::Start(element) => {
                if depth == 1 {
                    child_start = before;
                    if element.name().as_ref() == b"is" {
                        return Ok(None);
                    }
                }
                depth += 1;
            }
            Event::Empty(element) if depth == 1 => match element.name().as_ref() {
                b"f" if formula_end.replace(after).is_some() => return Ok(None),
                b"v" if value.replace(before..after).is_some() => return Ok(None),
                b"is" => return Ok(None),
                _ => {}
            },
            Event::End(element) => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
                if depth == 1 {
                    match element.name().as_ref() {
                        b"f" if formula_end.replace(after).is_some() => return Ok(None),
                        b"v" if value.replace(child_start..after).is_some() => return Ok(None),
                        _ => {}
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(formula_end.and_then(|end| match value {
        Some(value) if value.start >= end => Some(value),
        None => Some(end..end),
        _ => None,
    }))
}

fn set_row_height(
    attributes: &mut Vec<XmlAttribute>,
    height: Option<f64>,
    hidden_height: Option<f64>,
) {
    remove_attribute(attributes, "ht");
    remove_attribute(attributes, "customHeight");
    remove_attribute(attributes, "hidden");
    if let Some(height) = height {
        let hidden = height == 0.0;
        let value = fmt_num(if hidden {
            hidden_height.unwrap_or(height)
        } else {
            height
        });
        set_attribute(attributes, "ht", "ht", value);
        set_attribute(attributes, "customHeight", "customHeight", "1".to_owned());
        if hidden {
            set_attribute(attributes, "hidden", "hidden", "1".to_owned());
        }
    }
}

/// The style ascending, disjoint `runs` give `col`.
fn run_style(runs: &[ColStyle], col: ColId) -> Option<u32> {
    let at = runs.partition_point(|run| run.last < col);
    runs.get(at)
        .filter(|run| run.first <= col)
        .map(|run| run.xf)
}

fn set_col_width(
    attributes: &mut Vec<XmlAttribute>,
    width: Option<f64>,
    hidden_width: Option<f64>,
) {
    remove_attribute(attributes, "width");
    remove_attribute(attributes, "customWidth");
    remove_attribute(attributes, "hidden");
    remove_attribute(attributes, "bestFit");
    if let Some(width) = width {
        let hidden = width == 0.0;
        let value = fmt_num(if hidden {
            hidden_width.unwrap_or(width)
        } else {
            width
        });
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
        element.push_attribute(text_attribute(&attribute.name, &attribute.value));
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

/// First markup a model-only `<sheetData>` write would lose.
pub(crate) fn sheet_data_regeneration_loss(data: &[u8]) -> Result<Option<String>, ParseError> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().expand_empty_elements = false;
    reader.config_mut().check_end_names = true;
    let mut path: Vec<Vec<u8>> = Vec::new();
    let mut row_ordinal = 0;
    let mut cell_ordinal = 0;
    let mut row_location = "sheetData".to_owned();
    let mut cell_location = None;
    let mut cell_children = 0_u8;
    let mut inline_text = false;
    let loss = |markup: String, location: &str| {
        Ok(Some(format!(
            "regenerating sheetData would drop {markup} at {location}"
        )))
    };
    loop {
        let event = reader.read_event().map_err(xml_err)?;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let name = element.local_name().as_ref().to_vec();
                let parent = path.last().map(Vec::as_slice);
                if parent == Some(b"sheetData") && name == b"row" {
                    row_ordinal += 1;
                    cell_ordinal = 0;
                    row_location = format!(
                        "row {}",
                        attr(&element, b"r")?.unwrap_or_else(|| row_ordinal.to_string())
                    );
                } else if parent == Some(b"row") && name == b"c" {
                    cell_ordinal += 1;
                    cell_location = Some(attr(&element, b"r")?.map_or_else(
                        || format!("{row_location}, cell {cell_ordinal}"),
                        |address| format!("cell {address}"),
                    ));
                    cell_children = 0;
                }
                let location = cell_location.as_deref().unwrap_or(&row_location);
                let allowed = match parent {
                    None => name == b"sheetData",
                    Some(b"sheetData") => name == b"row",
                    Some(b"row") => name == b"c",
                    Some(b"c") => {
                        let bit = match name.as_slice() {
                            b"f" => 1,
                            b"v" => 2,
                            b"is" => 4,
                            _ => 0,
                        };
                        let allowed = bit != 0
                            && cell_children & bit == 0
                            && (bit & 6 == 0 || cell_children & 6 == 0);
                        cell_children |= bit;
                        if name == b"is" {
                            inline_text = false;
                        }
                        allowed
                    }
                    Some(b"is") => {
                        let allowed = name == b"t" && !inline_text;
                        inline_text = true;
                        allowed
                    }
                    _ => false,
                };
                if !allowed {
                    return loss(format!("<{}>", String::from_utf8_lossy(&name)), location);
                }
                let row_format = name == b"row"
                    && attr(&element, b"customFormat")?.is_some_and(|value| is_truthy(&value));
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(xml_err)?;
                    let key = attribute.key.as_ref();
                    if key == b"xmlns" || key.starts_with(b"xmlns:") {
                        continue;
                    }
                    let value = if matches!(
                        (name.as_slice(), key),
                        (b"c" | b"f", b"t") | (b"t", b"xml:space")
                    ) {
                        attribute
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .map_err(xml_err)?
                    } else {
                        String::from_utf8_lossy(attribute.value.as_ref())
                    };
                    let allowed = match name.as_slice() {
                        b"row" if key == b"s" => row_format && value.parse::<u32>().is_ok(),
                        b"row" => {
                            matches!(
                                key,
                                b"r" | b"ht"
                                    | b"customFormat"
                                    | b"customHeight"
                                    | b"hidden"
                                    | b"spans"
                                    | b"thickTop"
                                    | b"thickBot"
                            ) || attribute.key.local_name().as_ref() == b"dyDescent"
                        }
                        b"c" => match key {
                            b"r" | b"s" => true,
                            b"t" => matches!(
                                value.as_ref(),
                                "n" | "s" | "str" | "inlineStr" | "b" | "e"
                            ),
                            _ => false,
                        },
                        b"f" => match key {
                            b"t" => matches!(value.as_ref(), "shared" | "array" | "normal"),
                            b"ref" | b"si" => true,
                            _ => false,
                        },
                        b"t" => {
                            key == b"xml:space" && matches!(value.as_ref(), "preserve" | "default")
                        }
                        _ => false,
                    };
                    if !allowed {
                        return loss(
                            format!(
                                "<{}> attribute {}=\"{value}\"",
                                String::from_utf8_lossy(&name),
                                String::from_utf8_lossy(key),
                            ),
                            location,
                        );
                    }
                }
                if empty && name == b"is" {
                    return loss("<is> without a plain <t>".to_owned(), location);
                }
                if !empty {
                    path.push(name);
                } else if name == b"c" {
                    cell_location = None;
                } else if name == b"row" {
                    row_location = "sheetData".to_owned();
                }
            }
            Event::End(_) => {
                if path.last().is_some_and(|name| name == b"is") && !inline_text {
                    return loss(
                        "<is> without a plain <t>".to_owned(),
                        cell_location.as_deref().unwrap_or(&row_location),
                    );
                }
                match path.pop().as_deref() {
                    Some(b"c") => cell_location = None,
                    Some(b"row") => row_location = "sheetData".to_owned(),
                    _ => {}
                }
            }
            Event::Text(text)
                if !matches!(path.last().map(Vec::as_slice), Some(b"f" | b"v" | b"t")) =>
            {
                if !text.decode().map_err(xml_err)?.trim().is_empty() {
                    return loss(
                        "text".to_owned(),
                        cell_location.as_deref().unwrap_or(&row_location),
                    );
                }
            }
            Event::CData(_) | Event::GeneralRef(_)
                if matches!(path.last().map(Vec::as_slice), Some(b"f" | b"v" | b"t")) => {}
            Event::Text(_) => {}
            Event::Eof => return Ok(None),
            event => {
                let markup = match event {
                    Event::Comment(_) => "XML comment",
                    Event::PI(_) => "processing instruction",
                    Event::CData(_) => "CDATA outside cell text",
                    Event::GeneralRef(_) => "entity outside cell text",
                    _ => "XML markup",
                };
                return loss(
                    markup.to_owned(),
                    cell_location.as_deref().unwrap_or(&row_location),
                );
            }
        }
    }
}

/// Rows and cells of a `<sheetData>`; `None` when unpatchable cell by cell.
fn scan_sheet_data(data: &[u8]) -> Result<Option<SourceSheetData>, ParseError> {
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
    let mut alternates = Vec::new();
    let mut alternate: Option<(usize, Option<RangeInclusive<u32>>)> = None;
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(xml_err)?;
        let after = reader.buffer_position() as usize;
        if let Some((_, indices)) = &mut alternate
            && let Event::Start(element) | Event::Empty(element) = &event
        {
            let at = match (row.is_some(), element.local_name().as_ref()) {
                (true, b"c") => attr(element, b"r")?
                    .and_then(|address| CellRef::parse_a1(&address).ok())
                    .map(|at| at.col),
                (false, b"row") => attr(element, b"r")?
                    .and_then(|index| index.trim().parse::<u32>().ok())
                    .and_then(|index| index.checked_sub(1)),
                _ => None,
            };
            if let Some(at) = at {
                *indices = Some(indices.as_ref().map_or(at..=at, |indices| {
                    at.min(*indices.start())..=at.max(*indices.end())
                }));
            }
        }
        match event {
            Event::Start(element) => {
                if (depth == 1 || (depth == 2 && row.is_some()))
                    && alternate.is_none()
                    && element.local_name().as_ref() == b"AlternateContent"
                {
                    alternate = Some((depth, None));
                }
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
                            alternates: Vec::new(),
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
                    return Ok(Some(SourceSheetData {
                        element: expanded_empty(data, &element, before..after)?,
                        rows,
                        alternates,
                    }));
                }
                1 | 2 if element.local_name().as_ref() == b"AlternateContent" => {
                    return Ok(None);
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
                        alternates: Vec::new(),
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
                if let Some((_, indices)) = alternate.take_if(|(opened, _)| *opened == depth) {
                    let Some(indices) = indices else {
                        return Ok(None);
                    };
                    match &mut row {
                        Some(row) => row.alternates.push(SourceAlternate {
                            before: row.content_start(),
                            end: after,
                            indices,
                        }),
                        None => {
                            alternates.push(SourceAlternate {
                                before: cursor,
                                end: after,
                                indices,
                            });
                            cursor = after;
                        }
                    }
                }
                match depth {
                    0 => {
                        return Ok(Some(SourceSheetData {
                            element: SourceElement {
                                prefix,
                                suffix: data[before..after].to_vec(),
                                content: content_start..before,
                            },
                            rows,
                            alternates,
                        }));
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

pub(crate) fn patchable_alternate_grid(
    data: &[u8],
    wrapper_count: usize,
) -> Result<bool, ParseError> {
    let Some(grid) = scan_sheet_data(data)? else {
        return Ok(false);
    };
    let blocks = grid.alternates.len()
        + grid
            .rows
            .iter()
            .map(|row| row.alternates.len())
            .sum::<usize>();
    let ordered = |mut positions: Vec<_>| {
        positions.sort_unstable_by_key(|(offset, _)| *offset);
        positions
            .windows(2)
            .all(|pair: &[(usize, RangeInclusive<u32>)]| pair[0].1.end() < pair[1].1.start())
    };
    if blocks != wrapper_count
        || !ordered(
            grid.rows
                .iter()
                .map(|row| (row.span.start, row.index..=row.index))
                .chain(
                    grid.alternates
                        .iter()
                        .map(|block| (block.before, block.indices.clone())),
                )
                .collect(),
        )
    {
        return Ok(false);
    }
    Ok(grid
        .rows
        .iter()
        .filter(|row| !row.alternates.is_empty())
        .all(|row| {
            ordered(
                row.cells
                    .iter()
                    .map(|cell| (cell.span.start, cell.at.col..=cell.at.col))
                    .chain(
                        row.alternates
                            .iter()
                            .map(|block| (block.before, block.indices.clone())),
                    )
                    .collect(),
            )
        }))
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

#[cfg(test)]
#[test]
fn alternate_branches_stay_in_place_when_the_following_cell_is_removed() {
    let alternate = concat!(
        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:test:ext">"#,
        r#"<mc:Choice Requires="x"><c r="B1"><v>2</v></c></mc:Choice>"#,
        r#"<mc:Fallback><c r="B1"><v>3</v></c></mc:Fallback>"#,
        "</mc:AlternateContent>",
    );
    let source = format!(
        r#"<sheetData><row r="1"><c r="A1"><v>1</v></c>{alternate}<c r="D1"><v>4</v></c><c r="E1"><f>A1+D1</f><v>5</v></c></row></sheetData>"#,
    );
    let mut parts = crate::tests::package(&source, &[], false);
    let parsed = crate::tests::parse_workbook_with_package(&parts).unwrap();
    let facts = parsed.package.source_cell_facts(0).unwrap();
    assert!(!facts.non_inert_alternate);
    assert_eq!(facts.alternate_cells, BTreeSet::from([(0, 1)]));
    let mut workbook = parsed.workbook.clone();
    let row = &mut workbook.sheets[0];
    row.set_cell(CellRef::new(0, 3), Cell::default());
    row.cell_mut(CellRef::new(0, 4)).unwrap().value = CellValue::Number { value: 1.0 };

    let saved = style_match_tests::save(&parsed, &workbook, Some(SheetAxes::default()));
    assert_eq!(saved.matches(alternate).count(), 1, "{saved}");
    parts
        .iter_mut()
        .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
        .unwrap()
        .1 = saved.into_bytes();
    let reopened = crate::parse_workbook(&parts).unwrap();
    assert!(reopened.sheets[0].cell(CellRef::new(0, 3)).is_none());
    assert_eq!(
        reopened.sheets[0].cell(CellRef::new(0, 4)).unwrap().value,
        CellValue::Number { value: 1.0 },
    );
}

#[cfg(test)]
#[test]
fn alternate_grid_facts_record_both_ends_of_cell_and_row_spans() {
    let cells = r#"<c r="B1"><v>2</v></c><c r="D1"><v>4</v></c>"#;
    let rows = r#"<row r="2"><c r="B2"><v>2</v></c></row><row r="4"><c r="D4"><v>4</v></c></row>"#;
    let wrapper = |xml| {
        format!(
            r#"<mc:AlternateContent><mc:Choice Requires="x">{xml}</mc:Choice><mc:Fallback>{xml}</mc:Fallback></mc:AlternateContent>"#,
        )
    };
    let xml = format!(
        r#"<worksheet xmlns:mc="{}" xmlns:x="urn:test:ext"><sheetData><row r="1">{}</row>{}</sheetData></worksheet>"#,
        crate::drawing_anchors::NS_MARKUP_COMPATIBILITY,
        wrapper(cells),
        wrapper(rows),
    );
    let mut facts = SourceCellFacts::default();
    scan_alternate_content(xml.as_bytes(), true, &mut facts).unwrap();
    assert!(!facts.non_inert_alternate);
    assert_eq!(
        facts.alternate_spans,
        [
            AlternateSpan {
                rows: 0..1,
                cols: 1..4,
                wraps_rows: false
            },
            AlternateSpan {
                rows: 1..4,
                cols: 1..4,
                wraps_rows: true
            },
        ],
    );
}

#[cfg(test)]
#[test]
fn alternate_grid_facts_detect_empty_wrappers_and_unmodeled_formulas() {
    let cell = r#"<c r="A1"><v>2</v></c>"#;
    let choice = r#"<c r="A1"><f t="shared" si="0"/><v>2</v></c>"#;
    for (xml, formula) in [
        ("<mc:AlternateContent/>".to_owned(), false),
        (
            format!(
                r#"<mc:AlternateContent><mc:Choice Requires="x">{choice}</mc:Choice><mc:Fallback>{cell}</mc:Fallback></mc:AlternateContent>"#
            ),
            true,
        ),
    ] {
        let source = format!(
            r#"<worksheet xmlns:mc="{}" xmlns:x="urn:test:ext"><sheetData><row r="1">{xml}</row></sheetData></worksheet>"#,
            crate::drawing_anchors::NS_MARKUP_COMPATIBILITY,
        );
        let mut facts = SourceCellFacts::default();
        scan_alternate_content(source.as_bytes(), true, &mut facts).unwrap();
        assert!(facts.has_sheet_data_alternates());
        assert!(facts.non_inert_alternate);
        assert_eq!(facts.sheet_data_alternate_formulas, formula);
        assert!(!facts.alternate_formulas);
    }
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
