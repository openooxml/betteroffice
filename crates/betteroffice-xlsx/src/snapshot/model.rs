use std::collections::VecDeque;
use std::ops::Bound::{Excluded, Unbounded};

use xlsx_model::{
    Alignment, AnchorCell, AnchorEditAs, AnchorExtent, AnchorPos, Border, BorderEdge, BorderStyle,
    Cell, CellRef, ChartAnchor, ChartRef, ChartRefKind, ColStyle, Color, DateSystem, DefinedName,
    Fill, Font, FreezePane, HAlign, Hyperlink, Sheet, SheetChart, SheetFormat, SheetId, Table,
    Theme, VAlign, Workbook, Xf,
};

use super::cells::{CellCursor, read_cell, read_range, read_ref, write_range, write_ref};
use super::wire::{ChunkKind, Reader, Writer, frame, unframe};
use super::{SnapshotBudget, SnapshotError, SnapshotResult};

const HEADER: u8 = 0;
const NAME: u8 = 1;
const STRING: u8 = 2;
const FONT: u8 = 3;
const FILL: u8 = 4;
const BORDER: u8 = 5;
const XF: u8 = 6;
const NUM_FMT: u8 = 7;
const THEME: u8 = 8;
const INDEXED: u8 = 9;
const TABLE: u8 = 10;
const TABLE_COLUMN: u8 = 11;
const SHEET: u8 = 12;
const HYPERLINK: u8 = 13;
const MERGE: u8 = 14;
const COL_WIDTH: u8 = 15;
const ROW_HEIGHT: u8 = 16;
const COL_STYLE: u8 = 17;
const CHART: u8 = 18;
const CHART_REF: u8 = 19;
const ARRAY: u8 = 20;

pub(crate) struct ModelSnapshotEncoder {
    ordinal: u64,
    started: bool,
    metadata_done: bool,
    metadata: MetadataCursor,
    cells: CellCursor,
}

impl ModelSnapshotEncoder {
    pub(crate) fn new() -> Self {
        Self {
            ordinal: 0,
            started: false,
            metadata_done: false,
            metadata: MetadataCursor::default(),
            cells: CellCursor::default(),
        }
    }

    pub(crate) fn next(
        &mut self,
        model: &Workbook,
        budget: SnapshotBudget,
    ) -> SnapshotResult<Option<Vec<u8>>> {
        if !self.metadata_done {
            let mut payload = Writer::new();
            let mut count = 0;
            if !self.started {
                write_header(&mut payload, model)?;
                if frame(ChunkKind::Model, self.ordinal, &[]).len() + payload.len()
                    > budget.max_bytes()
                {
                    return Err(SnapshotError::new(
                        "snapshot model header exceeds byte budget",
                    ));
                }
                self.started = true;
                count = 1;
            }
            while count < budget.max_records() {
                let mut cursor = self.metadata;
                let mut record = Writer::new();
                if !cursor.write_next(&mut record, model)? {
                    self.metadata_done = true;
                    break;
                }
                let length =
                    frame(ChunkKind::Model, self.ordinal, &[]).len() + payload.len() + record.len();
                if length > budget.max_bytes() {
                    if count != 0 {
                        break;
                    }
                    return Err(SnapshotError::new(
                        "snapshot metadata record exceeds byte budget",
                    ));
                }
                payload.raw(&record.into_bytes());
                self.metadata = cursor;
                count += 1;
            }
            if !payload.is_empty() {
                let chunk = frame(ChunkKind::Model, self.ordinal, &payload.into_bytes());
                self.ordinal += 1;
                return Ok(Some(chunk));
            }
        }
        let chunk = self.cells.next(model, budget, self.ordinal)?;
        if chunk.is_some() {
            self.ordinal += 1;
        }
        Ok(chunk)
    }
}

fn exact_len(iter: impl Iterator) -> SnapshotResult<usize> {
    let (lower, upper) = iter.size_hint();
    if upper == Some(lower) {
        Ok(lower)
    } else {
        Err(SnapshotError::new(
            "snapshot model iterator has no exact size",
        ))
    }
}

fn global_counts(model: &Workbook) -> [usize; 10] {
    let Workbook {
        sheets,
        date_system: _,
        defined_names,
        shared_strings,
        styles,
        tables,
    } = model;
    [
        defined_names.len(),
        shared_strings.len(),
        styles.fonts.len(),
        styles.fills.len(),
        styles.borders.len(),
        styles.cell_xfs.len(),
        styles.num_fmts.len(),
        styles.indexed_colors.len(),
        tables.len(),
        sheets.len(),
    ]
}

fn sheet_counts(sheet: &Sheet) -> SnapshotResult<[usize; 8]> {
    Ok([
        sheet.hyperlinks.len(),
        sheet.merges.len(),
        sheet.col_widths.len(),
        sheet.row_heights.len(),
        sheet.col_styles.len(),
        sheet.charts.len(),
        exact_len(sheet.array_formulas())?,
        exact_len(sheet.iter_cells())?,
    ])
}

fn add_count(total: &mut usize, value: usize) -> SnapshotResult<()> {
    *total = total
        .checked_add(value)
        .ok_or_else(|| SnapshotError::new("snapshot model count overflows usize"))?;
    Ok(())
}

fn write_header(w: &mut Writer, model: &Workbook) -> SnapshotResult<()> {
    let Workbook {
        sheets,
        date_system,
        defined_names: _,
        shared_strings: _,
        styles: _,
        tables,
    } = model;
    let counts = global_counts(model);
    let mut records = 13;
    for count in counts {
        add_count(&mut records, count)?;
    }
    for table in tables {
        add_count(&mut records, table.columns.len())?;
    }
    let mut cells = 0;
    for sheet in sheets {
        let counts = sheet_counts(sheet)?;
        for count in &counts[..7] {
            add_count(&mut records, *count)?;
        }
        add_count(&mut cells, counts[7])?;
        for chart in &sheet.charts {
            add_count(&mut records, chart.refs.len())?;
        }
    }
    w.u8(HEADER);
    w.u8(match date_system {
        DateSystem::V1900 => 0,
        DateSystem::V1904 => 1,
    });
    w.var_usize(records);
    w.var_usize(cells);
    for count in counts {
        w.var_usize(count);
    }
    Ok(())
}

#[derive(Clone, Copy)]
struct MetadataCursor {
    phase: u8,
    index: usize,
    child: usize,
    sheet: usize,
    sheet_phase: u8,
    map_key: Option<u32>,
}

impl Default for MetadataCursor {
    fn default() -> Self {
        Self {
            phase: NAME,
            index: 0,
            child: 0,
            sheet: 0,
            sheet_phase: SHEET,
            map_key: None,
        }
    }
}

impl MetadataCursor {
    fn write_next(&mut self, w: &mut Writer, model: &Workbook) -> SnapshotResult<bool> {
        loop {
            if self.phase == SHEET {
                let Some(sheet) = model.sheets.get(self.sheet) else {
                    return Ok(false);
                };
                if self.write_sheet(w, sheet)? {
                    return Ok(true);
                }
                self.sheet += 1;
                self.sheet_phase = SHEET;
                continue;
            }
            let length = match self.phase {
                NAME => model.defined_names.len(),
                STRING => model.shared_strings.len(),
                FONT => model.styles.fonts.len(),
                FILL => model.styles.fills.len(),
                BORDER => model.styles.borders.len(),
                XF => model.styles.cell_xfs.len(),
                NUM_FMT => model.styles.num_fmts.len(),
                THEME => 12,
                INDEXED => model.styles.indexed_colors.len(),
                TABLE => model.tables.len(),
                _ => unreachable!(),
            };
            if self.index == length {
                self.phase = if self.phase == TABLE {
                    SHEET
                } else {
                    self.phase + 1
                };
                self.index = 0;
                continue;
            }
            let tag = if self.phase == TABLE && self.child != 0 {
                TABLE_COLUMN
            } else {
                self.phase
            };
            w.u8(tag);
            match self.phase {
                NAME => write_name(w, &model.defined_names[self.index]),
                STRING => w.str(&model.shared_strings[self.index]),
                FONT => write_font(w, &model.styles.fonts[self.index]),
                FILL => write_fill(w, &model.styles.fills[self.index]),
                BORDER => write_border(w, &model.styles.borders[self.index]),
                XF => write_xf(w, &model.styles.cell_xfs[self.index]),
                NUM_FMT => {
                    let (id, code) = &model.styles.num_fmts[self.index];
                    w.var_u32(u32::from(*id));
                    w.str(code);
                }
                THEME => {
                    let Theme { colors } = &model.styles.theme;
                    w.str(&colors[self.index]);
                }
                INDEXED => w.str(&model.styles.indexed_colors[self.index]),
                TABLE => {
                    let table = &model.tables[self.index];
                    if self.child == 0 {
                        write_table(w, table);
                    } else {
                        w.str(&table.columns[self.child - 1]);
                    }
                    self.child += 1;
                    if self.child <= table.columns.len() {
                        return Ok(true);
                    }
                    self.child = 0;
                }
                _ => unreachable!(),
            }
            self.index += 1;
            return Ok(true);
        }
    }

    fn write_sheet(&mut self, w: &mut Writer, sheet: &Sheet) -> SnapshotResult<bool> {
        loop {
            let length = match self.sheet_phase {
                SHEET => 1,
                HYPERLINK => sheet.hyperlinks.len(),
                MERGE => sheet.merges.len(),
                COL_WIDTH => sheet.col_widths.len(),
                ROW_HEIGHT => sheet.row_heights.len(),
                COL_STYLE => sheet.col_styles.len(),
                CHART => sheet.charts.len(),
                ARRAY => exact_len(sheet.array_formulas())?,
                _ => return Ok(false),
            };
            if self.index == length {
                self.sheet_phase = if self.sheet_phase == CHART {
                    ARRAY
                } else {
                    self.sheet_phase + 1
                };
                self.index = 0;
                self.map_key = None;
                continue;
            }
            w.u8(if self.sheet_phase == CHART && self.child != 0 {
                CHART_REF
            } else {
                self.sheet_phase
            });
            match self.sheet_phase {
                SHEET => {
                    w.str(&sheet.name);
                    w.option(sheet.freeze_pane.as_ref(), write_pane);
                    write_format(w, &sheet.format);
                    for count in sheet_counts(sheet)? {
                        w.var_usize(count);
                    }
                }
                HYPERLINK => write_hyperlink(w, &sheet.hyperlinks[self.index]),
                MERGE => write_range(w, &sheet.merges[self.index]),
                COL_WIDTH | ROW_HEIGHT => {
                    let map = if self.sheet_phase == COL_WIDTH {
                        &sheet.col_widths
                    } else {
                        &sheet.row_heights
                    };
                    let lower = self.map_key.map_or(Unbounded, Excluded);
                    let (&key, &value) = map
                        .range((lower, Unbounded))
                        .next()
                        .ok_or_else(|| SnapshotError::new("snapshot dimension cursor is stale"))?;
                    w.var_u32(key);
                    w.f64(value);
                    self.map_key = Some(key);
                }
                COL_STYLE => {
                    let ColStyle { first, last, xf } = &sheet.col_styles[self.index];
                    w.var_u32(*first);
                    w.var_u32(*last);
                    w.var_u32(*xf);
                }
                CHART => {
                    let chart = &sheet.charts[self.index];
                    if self.child == 0 {
                        write_chart(w, chart);
                    } else {
                        write_chart_ref(w, &chart.refs[self.child - 1]);
                    }
                    self.child += 1;
                    if self.child <= chart.refs.len() {
                        return Ok(true);
                    }
                    self.child = 0;
                }
                ARRAY => {
                    let (at, range) = sheet
                        .array_formulas()
                        .nth(self.index)
                        .ok_or_else(|| SnapshotError::new("snapshot array cursor is stale"))?;
                    write_ref(w, &at);
                    write_range(w, &range);
                }
                _ => unreachable!(),
            }
            self.index += 1;
            return Ok(true);
        }
    }
}

pub(crate) struct ModelSnapshotBuilder {
    model: Workbook,
    ordinal: u64,
    started: bool,
    failed: bool,
    remaining_records: usize,
    remaining_cells: usize,
    declared_cells: usize,
    expected_sheets: usize,
    runs: VecDeque<(u8, usize)>,
    sheet_cells: VecDeque<(usize, usize)>,
    cell_cursor: CellCursor,
    theme_index: usize,
    array_key: Option<(u32, u32)>,
}

impl ModelSnapshotBuilder {
    pub(crate) fn new() -> Self {
        Self {
            model: Workbook::default(),
            ordinal: 0,
            started: false,
            failed: false,
            remaining_records: 0,
            remaining_cells: 0,
            declared_cells: 0,
            expected_sheets: 0,
            runs: VecDeque::new(),
            sheet_cells: VecDeque::new(),
            cell_cursor: CellCursor::default(),
            theme_index: 0,
            array_key: None,
        }
    }

    pub(crate) fn push(&mut self, payload: &[u8]) -> SnapshotResult<()> {
        if self.failed {
            return Err(SnapshotError::new("snapshot model builder has failed"));
        }
        let result = self.push_inner(payload);
        if result.is_err() {
            self.failed = true;
        }
        result
    }

    fn push_inner(&mut self, chunk: &[u8]) -> SnapshotResult<()> {
        let (kind, ordinal, payload) = unframe(chunk)?;
        if ordinal != self.ordinal || payload.is_empty() {
            return Err(SnapshotError::new(
                "snapshot model chunk is missing or reordered",
            ));
        }
        let mut r = Reader::new(payload);
        match kind {
            ChunkKind::Model => {
                if self.started && self.remaining_records == 0 {
                    return Err(SnapshotError::new("unexpected snapshot model metadata"));
                }
                while !r.is_empty() {
                    let tag = r.u8()?;
                    if !self.started {
                        if tag != HEADER {
                            return Err(SnapshotError::new("snapshot model header is missing"));
                        }
                        self.read_header(&mut r)?;
                    } else {
                        self.read_record(tag, &mut r)?;
                    }
                }
            }
            ChunkKind::Cells => {
                if !self.started || self.remaining_records != 0 || !self.runs.is_empty() {
                    return Err(SnapshotError::new(
                        "snapshot cells precede complete metadata",
                    ));
                }
                self.read_cells(&mut r)?;
            }
            _ => return Err(SnapshotError::new("wrong snapshot model chunk kind")),
        }
        r.finish()?;
        self.ordinal += 1;
        Ok(())
    }

    fn read_header(&mut self, r: &mut Reader<'_>) -> SnapshotResult<()> {
        self.model.date_system = match r.u8()? {
            0 => DateSystem::V1900,
            1 => DateSystem::V1904,
            _ => return Err(SnapshotError::new("invalid snapshot date system")),
        };
        self.remaining_records = r
            .var_usize()?
            .checked_sub(1)
            .ok_or_else(|| SnapshotError::new("snapshot metadata count is zero"))?;
        self.remaining_cells = r.var_usize()?;
        let mut counts = [0; 10];
        for count in &mut counts {
            *count = r.var_usize()?;
        }
        self.expected_sheets = counts[9];
        for (tag, count) in [
            (NAME, counts[0]),
            (STRING, counts[1]),
            (FONT, counts[2]),
            (FILL, counts[3]),
            (BORDER, counts[4]),
            (XF, counts[5]),
            (NUM_FMT, counts[6]),
            (THEME, 12),
            (INDEXED, counts[7]),
            (TABLE, counts[8]),
            (SHEET, counts[9]),
        ] {
            if count != 0 {
                self.runs.push_back((tag, count));
            }
        }
        self.started = true;
        Ok(())
    }

    fn read_record(&mut self, tag: u8, r: &mut Reader<'_>) -> SnapshotResult<()> {
        let Some((expected, remaining)) = self.runs.front_mut() else {
            return Err(SnapshotError::new("extra snapshot metadata record"));
        };
        if tag != *expected || self.remaining_records == 0 {
            return Err(SnapshotError::new(
                "snapshot metadata records are reordered",
            ));
        }
        *remaining -= 1;
        if *remaining == 0 {
            self.runs.pop_front();
        }
        self.remaining_records -= 1;
        match tag {
            NAME => self.model.defined_names.push(read_name(r)?),
            STRING => self.model.shared_strings.push(r.str()?.to_owned()),
            FONT => self.model.styles.fonts.push(read_font(r)?),
            FILL => self.model.styles.fills.push(read_fill(r)?),
            BORDER => self.model.styles.borders.push(read_border(r)?),
            XF => self.model.styles.cell_xfs.push(read_xf(r)?),
            NUM_FMT => self
                .model
                .styles
                .num_fmts
                .push((read_u16(r)?, r.str()?.to_owned())),
            THEME => {
                self.model.styles.theme.colors[self.theme_index] = r.str()?.to_owned();
                self.theme_index += 1;
            }
            INDEXED => self.model.styles.indexed_colors.push(r.str()?.to_owned()),
            TABLE => {
                let (table, columns) = read_table(r)?;
                self.model.tables.push(table);
                if columns != 0 {
                    self.runs.push_front((TABLE_COLUMN, columns));
                }
            }
            TABLE_COLUMN => self
                .model
                .tables
                .last_mut()
                .ok_or_else(|| SnapshotError::new("snapshot table column has no table"))?
                .columns
                .push(r.str()?.to_owned()),
            SHEET => self.read_sheet(r)?,
            _ => self.read_sheet_record(tag, r)?,
        }
        if self.remaining_records == 0 {
            if !self.runs.is_empty()
                || self.model.sheets.len() != self.expected_sheets
                || self.declared_cells != self.remaining_cells
            {
                return Err(SnapshotError::new("snapshot model counts do not match"));
            }
        } else if self.runs.is_empty() {
            return Err(SnapshotError::new(
                "snapshot model has excess metadata count",
            ));
        }
        Ok(())
    }

    fn read_sheet(&mut self, r: &mut Reader<'_>) -> SnapshotResult<()> {
        let mut sheet = Sheet::new(r.str()?);
        sheet.freeze_pane = r.option(read_pane)?;
        sheet.format = read_format(r)?;
        let mut counts = [0; 8];
        for count in &mut counts {
            *count = r.var_usize()?;
        }
        add_count(&mut self.declared_cells, counts[7])?;
        if counts[7] != 0 {
            self.sheet_cells
                .push_back((self.model.sheets.len(), counts[7]));
        }
        self.model.sheets.push(sheet);
        self.array_key = None;
        for (tag, count) in [
            (ARRAY, counts[6]),
            (CHART, counts[5]),
            (COL_STYLE, counts[4]),
            (ROW_HEIGHT, counts[3]),
            (COL_WIDTH, counts[2]),
            (MERGE, counts[1]),
            (HYPERLINK, counts[0]),
        ] {
            if count != 0 {
                self.runs.push_front((tag, count));
            }
        }
        Ok(())
    }

    fn read_sheet_record(&mut self, tag: u8, r: &mut Reader<'_>) -> SnapshotResult<()> {
        let sheet = self
            .model
            .sheets
            .last_mut()
            .ok_or_else(|| SnapshotError::new("snapshot sheet metadata has no sheet"))?;
        match tag {
            HYPERLINK => sheet.hyperlinks.push(read_hyperlink(r)?),
            MERGE => sheet.merges.push(read_range(r)?),
            COL_WIDTH | ROW_HEIGHT => {
                let key = r.var_u32()?;
                let value = r.f64()?;
                let map = if tag == COL_WIDTH {
                    &mut sheet.col_widths
                } else {
                    &mut sheet.row_heights
                };
                if map
                    .last_key_value()
                    .is_some_and(|(&previous, _)| key <= previous)
                {
                    return Err(SnapshotError::new("snapshot dimensions are reordered"));
                }
                map.insert(key, value);
            }
            COL_STYLE => sheet.col_styles.push(ColStyle {
                first: r.var_u32()?,
                last: r.var_u32()?,
                xf: r.var_u32()?,
            }),
            CHART => {
                let (chart, refs) = read_chart(r)?;
                sheet.charts.push(chart);
                if refs != 0 {
                    self.runs.push_front((CHART_REF, refs));
                }
            }
            CHART_REF => sheet
                .charts
                .last_mut()
                .ok_or_else(|| SnapshotError::new("snapshot chart reference has no chart"))?
                .refs
                .push(read_chart_ref(r)?),
            ARRAY => {
                let at = read_ref(r)?;
                if at.abs_row || at.abs_col {
                    return Err(SnapshotError::new("snapshot array key has absolute flags"));
                }
                let key = (at.row, at.col);
                if self.array_key.is_some_and(|previous| key <= previous) {
                    return Err(SnapshotError::new("snapshot array keys are reordered"));
                }
                sheet.set_array_formula(at, read_range(r)?);
                self.array_key = Some(key);
            }
            _ => return Err(SnapshotError::new("invalid snapshot metadata tag")),
        }
        Ok(())
    }

    fn read_cells(&mut self, r: &mut Reader<'_>) -> SnapshotResult<()> {
        let sheet_index = r.var_usize()?;
        let base = (r.var_u32()?, r.var_u32()?);
        let count = r.var_usize()?;
        let Some(&(expected_sheet, remaining)) = self.sheet_cells.front() else {
            return Err(SnapshotError::new("extra snapshot cell block"));
        };
        if sheet_index != expected_sheet || count == 0 || count > remaining {
            return Err(SnapshotError::new(
                "snapshot cell block is missing or reordered",
            ));
        }
        if sheet_index != self.cell_cursor.sheet {
            self.cell_cursor.sheet = sheet_index;
            self.cell_cursor.after = None;
        }
        if base != self.cell_cursor.after.unwrap_or((0, 0)) {
            return Err(SnapshotError::new(
                "snapshot cell block is missing or reordered",
            ));
        }
        let sheet = &mut self.model.sheets[sheet_index];
        let mut previous = base;
        for _ in 0..count {
            let row_delta = r.var_u32()?;
            let col_delta = r.var_u32()?;
            let row = previous
                .0
                .checked_add(row_delta)
                .ok_or_else(|| SnapshotError::new("snapshot cell row overflows u32"))?;
            let col = if row_delta == 0 {
                previous
                    .1
                    .checked_add(col_delta)
                    .ok_or_else(|| SnapshotError::new("snapshot cell column overflows u32"))?
            } else {
                col_delta
            };
            let key = (row, col);
            if self.cell_cursor.after.is_some_and(|after| key <= after) {
                return Err(SnapshotError::new("snapshot cells are reordered"));
            }
            let at = CellRef::new(row, col);
            let cell = read_cell(r)?;
            if cell == Cell::default() {
                sheet.set_cell(
                    at,
                    Cell {
                        value: xlsx_model::CellValue::Empty,
                        formula: None,
                        style: Some(0),
                    },
                );
                *sheet.cell_mut(at).expect("inserted cell") = cell;
            } else {
                sheet.set_cell(at, cell);
            }
            self.cell_cursor.after = Some(key);
            previous = key;
        }
        if count == remaining {
            self.sheet_cells.pop_front();
        } else {
            self.sheet_cells.front_mut().expect("current cell sheet").1 -= count;
        }
        self.remaining_cells -= count;
        Ok(())
    }

    pub(crate) fn finish(self) -> SnapshotResult<Workbook> {
        if self.failed
            || !self.started
            || self.remaining_records != 0
            || self.remaining_cells != 0
            || !self.runs.is_empty()
            || !self.sheet_cells.is_empty()
        {
            return Err(SnapshotError::new("snapshot model is incomplete"));
        }
        Ok(self.model)
    }
}

fn read_u16(r: &mut Reader<'_>) -> SnapshotResult<u16> {
    u16::try_from(r.var_u32()?).map_err(|_| SnapshotError::new("snapshot integer overflows u16"))
}

fn write_name(w: &mut Writer, value: &DefinedName) {
    let DefinedName {
        name,
        formula,
        local_sheet,
        hidden,
    } = value;
    w.str(name);
    w.str(formula);
    w.option(*local_sheet, |w, sheet| {
        let SheetId(index) = sheet;
        w.var_u32(index);
    });
    w.bool(*hidden);
}

fn read_name(r: &mut Reader<'_>) -> SnapshotResult<DefinedName> {
    Ok(DefinedName {
        name: r.str()?.to_owned(),
        formula: r.str()?.to_owned(),
        local_sheet: r.option(|r| Ok(SheetId(r.var_u32()?)))?,
        hidden: r.bool()?,
    })
}

fn write_table(w: &mut Writer, value: &Table) {
    let Table {
        name,
        sheet,
        range,
        header_rows,
        totals_rows,
        columns,
    } = value;
    let SheetId(index) = sheet;
    w.str(name);
    w.var_u32(*index);
    write_range(w, range);
    w.var_u32(*header_rows);
    w.var_u32(*totals_rows);
    w.var_usize(columns.len());
}

fn read_table(r: &mut Reader<'_>) -> SnapshotResult<(Table, usize)> {
    let table = Table {
        name: r.str()?.to_owned(),
        sheet: SheetId(r.var_u32()?),
        range: read_range(r)?,
        header_rows: r.var_u32()?,
        totals_rows: r.var_u32()?,
        columns: Vec::new(),
    };
    Ok((table, r.var_usize()?))
}

fn write_pane(w: &mut Writer, value: &FreezePane) {
    let FreezePane {
        rows,
        cols,
        top_left,
    } = value;
    w.var_u32(*rows);
    w.var_u32(*cols);
    write_ref(w, top_left);
}

fn read_pane(r: &mut Reader<'_>) -> SnapshotResult<FreezePane> {
    Ok(FreezePane {
        rows: r.var_u32()?,
        cols: r.var_u32()?,
        top_left: read_ref(r)?,
    })
}

fn write_format(w: &mut Writer, value: &SheetFormat) {
    let SheetFormat {
        default_row_height_pt,
        custom_height,
        zero_height,
    } = value;
    w.option(*default_row_height_pt, Writer::f64);
    w.bool(*custom_height);
    w.bool(*zero_height);
}

fn read_format(r: &mut Reader<'_>) -> SnapshotResult<SheetFormat> {
    Ok(SheetFormat {
        default_row_height_pt: r.option(Reader::f64)?,
        custom_height: r.bool()?,
        zero_height: r.bool()?,
    })
}

fn write_hyperlink(w: &mut Writer, value: &Hyperlink) {
    let Hyperlink {
        range,
        external_target,
        location,
        tooltip,
        display,
    } = value;
    write_range(w, range);
    for value in [external_target, location, tooltip, display] {
        w.option(value.as_deref(), Writer::str);
    }
}

fn read_hyperlink(r: &mut Reader<'_>) -> SnapshotResult<Hyperlink> {
    Ok(Hyperlink {
        range: read_range(r)?,
        external_target: r.option(|r| Ok(r.str()?.to_owned()))?,
        location: r.option(|r| Ok(r.str()?.to_owned()))?,
        tooltip: r.option(|r| Ok(r.str()?.to_owned()))?,
        display: r.option(|r| Ok(r.str()?.to_owned()))?,
    })
}

fn write_color(w: &mut Writer, value: &Color) {
    match value {
        Color::Rgb(value) => {
            w.u8(0);
            w.str(value);
        }
        Color::Theme { idx, tint } => {
            w.u8(1);
            w.u8(*idx);
            w.f64(*tint);
        }
        Color::Indexed(value) => {
            w.u8(2);
            w.u8(*value);
        }
        Color::Auto => w.u8(3),
    }
}

fn read_color(r: &mut Reader<'_>) -> SnapshotResult<Color> {
    Ok(match r.u8()? {
        0 => Color::Rgb(r.str()?.to_owned()),
        1 => Color::Theme {
            idx: r.u8()?,
            tint: r.f64()?,
        },
        2 => Color::Indexed(r.u8()?),
        3 => Color::Auto,
        _ => return Err(SnapshotError::new("invalid snapshot color")),
    })
}

fn write_font(w: &mut Writer, value: &Font) {
    let Font {
        name,
        size_pt,
        bold,
        italic,
        underline,
        strike,
        color,
    } = value;
    w.option(name.as_deref(), Writer::str);
    w.option(*size_pt, Writer::f64);
    w.bool(*bold);
    w.bool(*italic);
    w.bool(*underline);
    w.bool(*strike);
    w.option(color.as_ref(), write_color);
}

fn read_font(r: &mut Reader<'_>) -> SnapshotResult<Font> {
    Ok(Font {
        name: r.option(|r| Ok(r.str()?.to_owned()))?,
        size_pt: r.option(Reader::f64)?,
        bold: r.bool()?,
        italic: r.bool()?,
        underline: r.bool()?,
        strike: r.bool()?,
        color: r.option(read_color)?,
    })
}

fn write_fill(w: &mut Writer, value: &Fill) {
    match value {
        Fill::None => w.u8(0),
        Fill::Solid(color) => {
            w.u8(1);
            write_color(w, color);
        }
    }
}

fn read_fill(r: &mut Reader<'_>) -> SnapshotResult<Fill> {
    Ok(match r.u8()? {
        0 => Fill::None,
        1 => Fill::Solid(read_color(r)?),
        _ => return Err(SnapshotError::new("invalid snapshot fill")),
    })
}

macro_rules! enum_codec {
    ($write:ident, $read:ident, $ty:ident, {$($variant:ident => $tag:literal),+ $(,)?}) => {
        fn $write(w: &mut Writer, value: &$ty) {
            w.u8(match value {
                $($ty::$variant => $tag),+
            });
        }

        fn $read(r: &mut Reader<'_>) -> SnapshotResult<$ty> {
            Ok(match r.u8()? {
                $($tag => $ty::$variant),+,
                _ => return Err(SnapshotError::new(concat!("invalid snapshot ", stringify!($ty)))),
            })
        }
    };
}

enum_codec!(write_border_style, read_border_style, BorderStyle, {
    Thin => 0,
    Medium => 1,
    Thick => 2,
    Dashed => 3,
    Dotted => 4,
    Double => 5,
    Hair => 6,
});

enum_codec!(write_halign, read_halign, HAlign, {
    General => 0,
    Left => 1,
    Center => 2,
    Right => 3,
    Fill => 4,
    Justify => 5,
    CenterContinuous => 6,
    Distributed => 7,
});

enum_codec!(write_valign, read_valign, VAlign, {
    Top => 0,
    Center => 1,
    Bottom => 2,
    Justify => 3,
    Distributed => 4,
});

enum_codec!(write_edit_as, read_edit_as, AnchorEditAs, {
    TwoCell => 0,
    OneCell => 1,
    Absolute => 2,
});

enum_codec!(write_ref_kind, read_ref_kind, ChartRefKind, {
    SeriesName => 0,
    Categories => 1,
    Values => 2,
    BubbleSize => 3,
    Title => 4,
    DataLabels => 5,
    Other => 6,
});

fn write_edge(w: &mut Writer, value: &BorderEdge) {
    let BorderEdge { style, color } = value;
    write_border_style(w, style);
    w.option(color.as_ref(), write_color);
}

fn read_edge(r: &mut Reader<'_>) -> SnapshotResult<BorderEdge> {
    Ok(BorderEdge {
        style: read_border_style(r)?,
        color: r.option(read_color)?,
    })
}

fn write_border(w: &mut Writer, value: &Border) {
    let Border {
        left,
        right,
        top,
        bottom,
    } = value;
    for edge in [left, right, top, bottom] {
        w.option(edge.as_ref(), write_edge);
    }
}

fn read_border(r: &mut Reader<'_>) -> SnapshotResult<Border> {
    Ok(Border {
        left: r.option(read_edge)?,
        right: r.option(read_edge)?,
        top: r.option(read_edge)?,
        bottom: r.option(read_edge)?,
    })
}

fn write_alignment(w: &mut Writer, value: &Alignment) {
    let Alignment {
        h,
        v,
        wrap_text,
        shrink_to_fit,
    } = value;
    w.option(h.as_ref(), write_halign);
    w.option(v.as_ref(), write_valign);
    w.bool(*wrap_text);
    w.bool(*shrink_to_fit);
}

fn read_alignment(r: &mut Reader<'_>) -> SnapshotResult<Alignment> {
    Ok(Alignment {
        h: r.option(read_halign)?,
        v: r.option(read_valign)?,
        wrap_text: r.bool()?,
        shrink_to_fit: r.bool()?,
    })
}

fn write_xf(w: &mut Writer, value: &Xf) {
    let Xf {
        font,
        fill,
        border,
        num_fmt_id,
        alignment,
    } = value;
    w.option(*font, Writer::var_u32);
    w.option(*fill, Writer::var_u32);
    w.option(*border, Writer::var_u32);
    w.option(*num_fmt_id, |w, id| w.var_u32(u32::from(id)));
    w.option(alignment.as_ref(), write_alignment);
}

fn read_xf(r: &mut Reader<'_>) -> SnapshotResult<Xf> {
    Ok(Xf {
        font: r.option(Reader::var_u32)?,
        fill: r.option(Reader::var_u32)?,
        border: r.option(Reader::var_u32)?,
        num_fmt_id: r.option(read_u16)?,
        alignment: r.option(read_alignment)?,
    })
}

fn write_anchor_cell(w: &mut Writer, value: &AnchorCell) {
    let AnchorCell {
        col,
        col_off,
        row,
        row_off,
    } = value;
    w.var_u32(*col);
    w.var_i64(*col_off);
    w.var_u32(*row);
    w.var_i64(*row_off);
}

fn read_anchor_cell(r: &mut Reader<'_>) -> SnapshotResult<AnchorCell> {
    Ok(AnchorCell {
        col: r.var_u32()?,
        col_off: r.var_i64()?,
        row: r.var_u32()?,
        row_off: r.var_i64()?,
    })
}

fn write_extent(w: &mut Writer, value: &AnchorExtent) {
    let AnchorExtent { cx, cy } = value;
    w.var_i64(*cx);
    w.var_i64(*cy);
}

fn read_extent(r: &mut Reader<'_>) -> SnapshotResult<AnchorExtent> {
    Ok(AnchorExtent {
        cx: r.var_i64()?,
        cy: r.var_i64()?,
    })
}

fn write_anchor(w: &mut Writer, value: &ChartAnchor) {
    match value {
        ChartAnchor::TwoCell { from, to, edit_as } => {
            w.u8(0);
            write_anchor_cell(w, from);
            write_anchor_cell(w, to);
            write_edit_as(w, edit_as);
        }
        ChartAnchor::OneCell { from, extent } => {
            w.u8(1);
            write_anchor_cell(w, from);
            write_extent(w, extent);
        }
        ChartAnchor::Absolute { pos, extent } => {
            let AnchorPos { x, y } = pos;
            w.u8(2);
            w.var_i64(*x);
            w.var_i64(*y);
            write_extent(w, extent);
        }
    }
}

fn read_anchor(r: &mut Reader<'_>) -> SnapshotResult<ChartAnchor> {
    Ok(match r.u8()? {
        0 => ChartAnchor::TwoCell {
            from: read_anchor_cell(r)?,
            to: read_anchor_cell(r)?,
            edit_as: read_edit_as(r)?,
        },
        1 => ChartAnchor::OneCell {
            from: read_anchor_cell(r)?,
            extent: read_extent(r)?,
        },
        2 => ChartAnchor::Absolute {
            pos: AnchorPos {
                x: r.var_i64()?,
                y: r.var_i64()?,
            },
            extent: read_extent(r)?,
        },
        _ => return Err(SnapshotError::new("invalid snapshot chart anchor")),
    })
}

fn write_chart(w: &mut Writer, value: &SheetChart) {
    let SheetChart {
        part,
        drawing,
        anchor_index,
        anchor,
        refs,
    } = value;
    w.str(part);
    w.str(drawing);
    w.var_usize(*anchor_index);
    write_anchor(w, anchor);
    w.var_usize(refs.len());
}

fn read_chart(r: &mut Reader<'_>) -> SnapshotResult<(SheetChart, usize)> {
    let chart = SheetChart {
        part: r.str()?.to_owned(),
        drawing: r.str()?.to_owned(),
        anchor_index: r.var_usize()?,
        anchor: read_anchor(r)?,
        refs: Vec::new(),
    };
    Ok((chart, r.var_usize()?))
}

fn write_chart_ref(w: &mut Writer, value: &ChartRef) {
    let ChartRef { kind, formula } = value;
    write_ref_kind(w, kind);
    w.str(formula);
}

fn read_chart_ref(r: &mut Reader<'_>) -> SnapshotResult<ChartRef> {
    Ok(ChartRef {
        kind: read_ref_kind(r)?,
        formula: r.str()?.to_owned(),
    })
}

#[cfg(test)]
mod tests {
    use xlsx_model::{CellRange, CellValue, ErrorValue, Stylesheet};

    use super::*;

    fn encode(model: &Workbook, budget: SnapshotBudget) -> Vec<Vec<u8>> {
        let mut encoder = ModelSnapshotEncoder::new();
        let mut chunks = Vec::new();
        while let Some(chunk) = encoder.next(model, budget).unwrap() {
            chunks.push(chunk);
        }
        chunks
    }

    fn decode(chunks: &[Vec<u8>]) -> Workbook {
        let mut builder = ModelSnapshotBuilder::new();
        for chunk in chunks {
            builder.push(chunk).unwrap();
        }
        builder.finish().unwrap()
    }

    fn sample_model() -> Workbook {
        let range = CellRange {
            start: CellRef {
                row: 7,
                col: 9,
                abs_row: true,
                abs_col: false,
            },
            end: CellRef {
                row: 2,
                col: 3,
                abs_row: false,
                abs_col: true,
            },
        };
        let mut styles = Stylesheet::default();
        styles.fonts.extend([
            Font {
                name: None,
                size_pt: None,
                bold: false,
                italic: false,
                underline: false,
                strike: false,
                color: None,
            },
            Font {
                name: Some(String::new()),
                size_pt: Some(-0.0),
                bold: true,
                italic: true,
                underline: true,
                strike: true,
                color: Some(Color::Theme { idx: 0, tint: -0.0 }),
            },
            Font {
                name: Some("Font".into()),
                size_pt: Some(11.25),
                bold: true,
                italic: false,
                underline: true,
                strike: false,
                color: Some(Color::Rgb(String::new())),
            },
        ]);
        styles.fills = vec![
            Fill::None,
            Fill::Solid(Color::Indexed(u8::MAX)),
            Fill::Solid(Color::Auto),
            Fill::Solid(Color::Rgb("#010203".into())),
            Fill::Solid(Color::Theme { idx: 9, tint: -0.5 }),
            Fill::None,
        ];
        for style in [
            BorderStyle::Thin,
            BorderStyle::Medium,
            BorderStyle::Thick,
            BorderStyle::Dashed,
            BorderStyle::Dotted,
            BorderStyle::Double,
            BorderStyle::Hair,
        ] {
            styles.borders.push(Border {
                left: Some(BorderEdge {
                    style,
                    color: Some(Color::Auto),
                }),
                right: None,
                top: Some(BorderEdge { style, color: None }),
                bottom: Some(BorderEdge {
                    style,
                    color: Some(Color::Indexed(0)),
                }),
            });
        }
        styles.cell_xfs.push(Xf::default());
        for h in [
            HAlign::General,
            HAlign::Left,
            HAlign::Center,
            HAlign::Right,
            HAlign::Fill,
            HAlign::Justify,
            HAlign::CenterContinuous,
            HAlign::Distributed,
        ] {
            for v in [
                VAlign::Top,
                VAlign::Center,
                VAlign::Bottom,
                VAlign::Justify,
                VAlign::Distributed,
            ] {
                styles.cell_xfs.push(Xf {
                    font: Some(0),
                    fill: Some(0),
                    border: Some(0),
                    num_fmt_id: Some(0),
                    alignment: Some(Alignment {
                        h: Some(h),
                        v: Some(v),
                        wrap_text: true,
                        shrink_to_fit: true,
                    }),
                });
            }
        }
        styles.cell_xfs.push(Xf {
            font: Some(u32::MAX),
            fill: None,
            border: None,
            num_fmt_id: Some(u16::MAX),
            alignment: Some(Alignment::default()),
        });
        styles.num_fmts = vec![
            (165, "code".into()),
            (164, String::new()),
            (165, "code".into()),
        ];
        styles.indexed_colors = vec![String::new(), "#123456".into(), String::new()];
        styles.theme.colors = std::array::from_fn(|index| format!("slot{index}"));
        let mut sheet = Sheet::new("");
        sheet.freeze_pane = Some(FreezePane {
            rows: 0,
            cols: u32::MAX,
            top_left: range.start,
        });
        sheet.hyperlinks = vec![
            Hyperlink {
                range,
                external_target: Some(String::new()),
                location: None,
                tooltip: Some("tip".into()),
                display: Some(String::new()),
            },
            Hyperlink {
                range,
                external_target: None,
                location: Some(String::new()),
                tooltip: None,
                display: None,
            },
        ];
        sheet.merges = vec![range, range];
        sheet.col_widths = [(9, 9.5), (0, -0.0), (u32::MAX, f64::INFINITY)]
            .into_iter()
            .collect();
        sheet.row_heights = [(u32::MAX, f64::NEG_INFINITY), (1, 0.0)]
            .into_iter()
            .collect();
        sheet.format = SheetFormat {
            default_row_height_pt: Some(-0.0),
            custom_height: true,
            zero_height: true,
        };
        sheet.col_styles = vec![
            ColStyle {
                first: 9,
                last: 3,
                xf: 0,
            },
            ColStyle {
                first: 0,
                last: 9,
                xf: u32::MAX,
            },
        ];
        let corner = AnchorCell {
            col: u32::MAX,
            col_off: i64::MIN,
            row: 0,
            row_off: i64::MAX,
        };
        let extent = AnchorExtent {
            cx: -1,
            cy: i64::MIN,
        };
        let anchors = [
            ChartAnchor::TwoCell {
                from: corner,
                to: corner,
                edit_as: AnchorEditAs::TwoCell,
            },
            ChartAnchor::TwoCell {
                from: corner,
                to: corner,
                edit_as: AnchorEditAs::OneCell,
            },
            ChartAnchor::TwoCell {
                from: corner,
                to: corner,
                edit_as: AnchorEditAs::Absolute,
            },
            ChartAnchor::OneCell {
                from: corner,
                extent,
            },
            ChartAnchor::Absolute {
                pos: AnchorPos { x: i64::MIN, y: -1 },
                extent,
            },
        ];
        for (anchor_index, anchor) in anchors.into_iter().enumerate() {
            sheet.charts.push(SheetChart {
                part: String::new(),
                drawing: "drawing".into(),
                anchor_index,
                anchor,
                refs: [
                    ChartRefKind::SeriesName,
                    ChartRefKind::Categories,
                    ChartRefKind::Values,
                    ChartRefKind::BubbleSize,
                    ChartRefKind::Title,
                    ChartRefKind::DataLabels,
                    ChartRefKind::Other,
                ]
                .into_iter()
                .map(|kind| ChartRef {
                    kind,
                    formula: String::new(),
                })
                .collect(),
            });
        }
        sheet.set_array_formula(CellRef::new(9, 2), range);
        sheet.set_array_formula(CellRef::new(1, 8), range);
        sheet.set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::Number { value: -0.0 },
                formula: Some(String::new()),
                style: Some(0),
            },
        );
        sheet.set_cell(
            CellRef::new(u32::MAX, u32::MAX),
            Cell {
                value: CellValue::Empty,
                formula: None,
                style: Some(0),
            },
        );
        Workbook {
            sheets: vec![sheet, Sheet::new(""), Sheet::new("Last")],
            date_system: DateSystem::V1904,
            defined_names: vec![
                DefinedName {
                    name: String::new(),
                    formula: String::new(),
                    local_sheet: None,
                    hidden: true,
                },
                DefinedName {
                    name: "duplicate".into(),
                    formula: "#REF!".into(),
                    local_sheet: Some(SheetId(0)),
                    hidden: false,
                },
            ],
            shared_strings: vec!["same".into(), String::new(), "same".into(), "水".into()],
            styles,
            tables: vec![
                Table {
                    name: String::new(),
                    sheet: SheetId(0),
                    range,
                    header_rows: 0,
                    totals_rows: u32::MAX,
                    columns: vec!["same".into(), String::new(), "same".into()],
                },
                Table {
                    name: "Empty".into(),
                    sheet: SheetId(u32::MAX),
                    range,
                    header_rows: u32::MAX,
                    totals_rows: 0,
                    columns: Vec::new(),
                },
            ],
        }
    }

    #[test]
    fn metadata_roundtrip_preserves_pool_order_and_style_presence() {
        let model = sample_model();
        let budget = SnapshotBudget::new(3, 120).unwrap();
        let chunks = encode(&model, budget);
        assert!(chunks.iter().all(|chunk| chunk.len() <= budget.max_bytes()));
        let decoded = decode(&chunks);
        assert_eq!(decoded, model);
        assert_eq!(decoded.shared_strings, ["same", "", "same", "水"]);
        assert_eq!(decoded.styles.cell_xfs[0].font, None);
        assert_eq!(decoded.styles.cell_xfs[1].font, Some(0));
        assert_eq!(decoded.styles.cell_xfs[0].num_fmt_id, None);
        assert_eq!(decoded.styles.cell_xfs[1].num_fmt_id, Some(0));
        assert_eq!(decoded.styles.fonts[0].name, None);
        assert_eq!(decoded.styles.fonts[1].name.as_deref(), Some(""));
        assert_eq!(
            decoded.sheets[0]
                .format
                .default_row_height_pt
                .unwrap()
                .to_bits(),
            1 << 63
        );
        assert_eq!(decoded.styles.fonts[1].size_pt.unwrap().to_bits(), 1 << 63);
        assert_eq!(encode(&decoded, budget), chunks);
    }

    #[test]
    fn row_continuations_preserve_cell_identity() {
        let mut sheet = Sheet::new("Dense");
        for col in 0..15 {
            sheet.set_cell(
                CellRef::new(0, col),
                Cell {
                    value: CellValue::Text {
                        value: format!("{col}"),
                    },
                    formula: if col % 2 == 0 {
                        Some(String::new())
                    } else {
                        None
                    },
                    style: if col % 3 == 0 { Some(0) } else { None },
                },
            );
        }
        *sheet.cell_mut(CellRef::new(0, 4)).unwrap() = Cell::default();
        sheet.set_cell(
            CellRef::new(200, 9),
            Cell {
                value: CellValue::Error {
                    value: ErrorValue::Spill,
                },
                formula: Some("A1".into()),
                style: None,
            },
        );
        let model = Workbook {
            sheets: vec![Sheet::new("Empty"), sheet, Sheet::new("Tail")],
            date_system: DateSystem::V1900,
            defined_names: Vec::new(),
            shared_strings: Vec::new(),
            styles: Stylesheet::default(),
            tables: Vec::new(),
        };
        for records in 1..=3 {
            let budget = SnapshotBudget::new(records, 80).unwrap();
            let chunks = encode(&model, budget);
            let mut emitted = 0;
            let mut continued = false;
            for chunk in &chunks {
                assert!(chunk.len() <= 80);
                let (kind, _, payload) = unframe(chunk).unwrap();
                if kind == ChunkKind::Cells {
                    let mut r = Reader::new(payload);
                    assert_eq!(r.var_usize().unwrap(), 1);
                    let row = r.var_u32().unwrap();
                    let col = r.var_u32().unwrap();
                    let count = r.var_usize().unwrap();
                    assert!(count <= records);
                    continued |= row == 0 && col != 0;
                    emitted += count;
                }
            }
            assert!(continued);
            assert_eq!(emitted, 16);
            let decoded = decode(&chunks);
            assert_eq!(decoded, model);
            assert_eq!(
                decoded.sheets[1].cell(CellRef::new(0, 4)),
                Some(&Cell::default())
            );
            assert_eq!(encode(&decoded, budget), chunks);
        }
    }

    #[test]
    fn cell_continuations_preserve_row_and_sheet_boundaries() {
        let mut sheet = Sheet::new("Boundary");
        for (index, (row, col)) in [
            (0, 3),
            (1, 0),
            (1, u32::MAX),
            (2, 0),
            (u32::MAX, 0),
            (u32::MAX, u32::MAX),
        ]
        .into_iter()
        .enumerate()
        {
            sheet.set_cell(
                CellRef::new(row, col),
                Cell {
                    value: CellValue::Text {
                        value: format!("cell-{index}"),
                    },
                    formula: None,
                    style: None,
                },
            );
        }
        let mut tail = Sheet::new("Tail");
        tail.set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::Bool { value: false },
                formula: Some(String::new()),
                style: Some(0),
            },
        );
        let model = Workbook {
            sheets: vec![Sheet::new("Empty"), sheet, Sheet::new("Gap"), tail],
            date_system: DateSystem::V1900,
            defined_names: Vec::new(),
            shared_strings: Vec::new(),
            styles: Stylesheet::default(),
            tables: Vec::new(),
        };
        for (records, bytes) in [(1, 80), (3, 80), (100, 48)] {
            let budget = SnapshotBudget::new(records, bytes).unwrap();
            let chunks = encode(&model, budget);
            assert!(chunks.iter().all(|chunk| chunk.len() <= bytes));
            let decoded = decode(&chunks);
            assert_eq!(decoded, model);
            assert_eq!(encode(&decoded, budget), chunks);
        }
    }

    #[test]
    fn current_spills_and_cached_values_survive_without_recalc() {
        let mut sheet = Sheet::new("Current");
        let spill = CellRange {
            start: CellRef::new(0, 0),
            end: CellRef {
                row: 1,
                col: 1,
                abs_row: true,
                abs_col: true,
            },
        };
        sheet.set_array_formula(CellRef::new(0, 0), spill);
        sheet.set_array_formula(CellRef::new(7, 8), spill);
        for (at, value, formula) in [
            (CellRef::new(0, 0), 900.0, Some("SEQUENCE(2,2)".into())),
            (CellRef::new(0, 1), -0.0, None),
            (CellRef::new(1, 0), 0.0, None),
            (
                CellRef::new(1, 1),
                f64::from_bits(0x7ff8_0000_0000_007b),
                None,
            ),
        ] {
            sheet.set_cell(
                at,
                Cell {
                    value: CellValue::Number { value },
                    formula,
                    style: None,
                },
            );
        }
        let mut model = Workbook::default();
        model.sheets.push(sheet);
        let budget = SnapshotBudget::new(1, 80).unwrap();
        let chunks = encode(&model, budget);
        let decoded = decode(&chunks);
        assert_eq!(
            decoded.sheets[0].array_formula(CellRef::new(0, 0)),
            Some(spill)
        );
        assert_eq!(
            decoded.sheets[0].array_formula(CellRef::new(7, 8)),
            Some(spill)
        );
        assert_eq!(
            decoded.sheets[0].cell(CellRef::new(0, 0)).unwrap().value,
            CellValue::Number { value: 900.0 }
        );
        for (at, bits) in [
            (CellRef::new(0, 1), (-0.0_f64).to_bits()),
            (CellRef::new(1, 0), 0.0_f64.to_bits()),
            (CellRef::new(1, 1), 0x7ff8_0000_0000_007b),
        ] {
            let CellValue::Number { value } = &decoded.sheets[0].cell(at).unwrap().value else {
                panic!("expected numeric cache");
            };
            assert_eq!(value.to_bits(), bits);
        }
        assert_eq!(encode(&decoded, budget), chunks);
    }

    #[test]
    fn model_reencode_is_byte_identical() {
        let mut model = sample_model();
        model.styles.fonts[1].size_pt = Some(f64::from_bits(0x7ff8_0000_0000_1234));
        model.styles.fonts[1].color = Some(Color::Theme {
            idx: 0,
            tint: f64::from_bits(0xfff8_0000_0000_5678),
        });
        model.sheets[0]
            .col_widths
            .insert(7, f64::from_bits(0x7ff0_0000_0000_0001));
        model.sheets[0].format.default_row_height_pt = Some(f64::from_bits(0x7ff8_0000_0000_0042));
        for col in 0..160 {
            model.sheets[2].set_cell(
                CellRef::new(3, col),
                Cell {
                    value: CellValue::Number {
                        value: f64::from_bits(0x7ff8_0000_0000_0000 | u64::from(col)),
                    },
                    formula: Some(String::new()),
                    style: Some(0),
                },
            );
        }
        for (records, bytes) in [(1, 120), (3, 120), (1_000, 512)] {
            let budget = SnapshotBudget::new(records, bytes).unwrap();
            let chunks = encode(&model, budget);
            assert_eq!(encode(&decode(&chunks), budget), chunks);
        }
        let model = Workbook::default();
        let budget = SnapshotBudget::new(1, 32).unwrap();
        let chunks = encode(&model, budget);
        assert_eq!(decode(&chunks), model);
        assert_eq!(encode(&decode(&chunks), budget), chunks);
    }

    #[test]
    fn builder_rejects_missing_or_reordered_chunks() {
        let chunks = encode(&sample_model(), SnapshotBudget::new(1, 120).unwrap());
        assert!(ModelSnapshotBuilder::new().finish().is_err());
        let mut builder = ModelSnapshotBuilder::new();
        builder.push(&chunks[0]).unwrap();
        assert!(builder.finish().is_err());
        let mut builder = ModelSnapshotBuilder::new();
        assert!(builder.push(&chunks[1]).is_err());
        let mut builder = ModelSnapshotBuilder::new();
        builder.push(&chunks[0]).unwrap();
        assert!(builder.push(&chunks[2]).is_err());
        assert!(builder.finish().is_err());
        let mut builder = ModelSnapshotBuilder::new();
        builder.push(&chunks[0]).unwrap();
        assert!(builder.push(&chunks[0]).is_err());
        let mut builder = ModelSnapshotBuilder::new();
        for chunk in &chunks[..chunks.len() - 1] {
            builder.push(chunk).unwrap();
        }
        assert!(builder.finish().is_err());
        let cell_chunk = chunks
            .iter()
            .find(|chunk| unframe(chunk).unwrap().0 == ChunkKind::Cells)
            .unwrap();
        let (_, _, payload) = unframe(cell_chunk).unwrap();
        let mut builder = ModelSnapshotBuilder::new();
        assert!(builder.push(&frame(ChunkKind::Cells, 0, payload)).is_err());
        let mut builder = ModelSnapshotBuilder::new();
        assert!(builder.push(&frame(ChunkKind::Preserved, 0, &[0])).is_err());
        let mut builder = ModelSnapshotBuilder::new();
        let (kind, ordinal, payload) = unframe(&chunks[0]).unwrap();
        let mut payload = payload.to_vec();
        payload.push(u8::MAX);
        assert!(builder.push(&frame(kind, ordinal, &payload)).is_err());
        let mut builder = ModelSnapshotBuilder::new();
        builder.push(&chunks[0]).unwrap();
        let mut truncated = chunks[1].clone();
        truncated.pop();
        assert!(builder.push(&truncated).is_err());
        assert!(builder.finish().is_err());
    }

    #[test]
    fn cell_byte_budget_admits_only_one_oversized_cell() {
        let mut model = Workbook::default();
        let mut sheet = Sheet::new("Bytes");
        for col in 0..3 {
            sheet.set_cell(
                CellRef::new(0, col),
                Cell {
                    value: CellValue::Text {
                        value: "x".repeat(400),
                    },
                    formula: None,
                    style: None,
                },
            );
        }
        model.sheets.push(sheet);
        let budget = SnapshotBudget::new(3, 80).unwrap();
        let chunks = encode(&model, budget);
        let mut cells = 0;
        for chunk in &chunks {
            let (kind, _, payload) = unframe(chunk).unwrap();
            if kind == ChunkKind::Cells {
                assert!(chunk.len() > 80);
                let mut r = Reader::new(payload);
                r.var_usize().unwrap();
                r.var_u32().unwrap();
                r.var_u32().unwrap();
                assert_eq!(r.var_usize().unwrap(), 1);
                cells += 1;
            }
        }
        assert_eq!(cells, 3);
        assert_eq!(decode(&chunks), model);
    }

    #[test]
    fn metadata_byte_budget_errors_leave_the_cursor_retryable() {
        let model = Workbook {
            sheets: Vec::new(),
            date_system: DateSystem::V1900,
            defined_names: Vec::new(),
            shared_strings: vec!["x".repeat(400)],
            styles: Stylesheet::default(),
            tables: Vec::new(),
        };
        let mut encoder = ModelSnapshotEncoder::new();
        assert!(
            encoder
                .next(&model, SnapshotBudget::new(3, 1).unwrap())
                .is_err()
        );
        let mut chunks = Vec::new();
        loop {
            match encoder.next(&model, SnapshotBudget::new(3, 80).unwrap()) {
                Ok(Some(chunk)) => {
                    assert!(chunk.len() <= 80);
                    chunks.push(chunk);
                }
                Err(_) => break,
                Ok(None) => panic!("oversized metadata record was accepted"),
            }
        }
        let budget = SnapshotBudget::new(3, 512).unwrap();
        while let Some(chunk) = encoder.next(&model, budget).unwrap() {
            assert!(chunk.len() <= 512);
            chunks.push(chunk);
        }
        assert_eq!(decode(&chunks), model);
    }

    #[test]
    fn freshly_opened_recalculated_model_roundtrips() {
        let mut source = Workbook::default();
        let mut sheet = Sheet::new("Open");
        sheet.set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::Number { value: 5.0 },
                formula: None,
                style: None,
            },
        );
        sheet.set_cell(
            CellRef::new(0, 1),
            Cell {
                value: CellValue::Number { value: -99.0 },
                formula: Some("A1+1".into()),
                style: None,
            },
        );
        source.sheets.push(sheet);
        let parts = xlsx_parse::serialize_workbook(&source).unwrap();
        let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
        let opened = crate::Workbook::open_recalculated(
            &bytes,
            crate::CalculationOptions {
                now_serial: Some(45_000.0),
            },
        )
        .unwrap();
        assert_eq!(
            opened.model().sheets[0]
                .cell(CellRef::new(0, 1))
                .unwrap()
                .value,
            CellValue::Number { value: 6.0 }
        );
        let budget = SnapshotBudget::new(2, 120).unwrap();
        let chunks = encode(opened.model(), budget);
        let decoded = decode(&chunks);
        assert_eq!(&decoded, opened.model());
        assert_eq!(encode(&decoded, budget), chunks);
    }
}
