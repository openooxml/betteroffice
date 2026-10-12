//! Sparse row and column extents.

#[cfg(feature = "test-counters")]
use std::cell::Cell;
use std::collections::{BTreeMap, HashSet};
use std::ops::Range;
use std::sync::Arc;

use xlsx_model::styles::{Font, Stylesheet};
use xlsx_model::workbook::Sheet;
use xlsx_model::{CellRef, ColId, MAX_COLS, MAX_ROWS, RowId};

use crate::Viewport;

#[cfg(feature = "test-counters")]
thread_local! {
    static GEOMETRY_CONSTRUCTIONS: Cell<u64> = const { Cell::new(0) };
    static AUTOFIT_CELL_VISITS: Cell<u64> = const { Cell::new(0) };
    static ROW_STORAGE_ALLOCATIONS: Cell<u64> = const { Cell::new(0) };
    static PREFIX_UPDATES: Cell<u64> = const { Cell::new(0) };
    static PREFIX_QUERIES: Cell<u64> = const { Cell::new(0) };
}

#[cfg(feature = "test-counters")]
#[doc(hidden)]
pub fn geometry_counters() -> (u64, u64) {
    (GEOMETRY_CONSTRUCTIONS.get(), AUTOFIT_CELL_VISITS.get())
}

#[cfg(feature = "test-counters")]
#[doc(hidden)]
pub fn reset_geometry_counters() {
    GEOMETRY_CONSTRUCTIONS.set(0);
    AUTOFIT_CELL_VISITS.set(0);
    ROW_STORAGE_ALLOCATIONS.set(0);
    PREFIX_UPDATES.set(0);
    PREFIX_QUERIES.set(0);
}

#[cfg(feature = "test-counters")]
#[doc(hidden)]
pub fn geometry_prefix_counters() -> (u64, u64) {
    (PREFIX_UPDATES.get(), PREFIX_QUERIES.get())
}

#[cfg(feature = "test-counters")]
#[doc(hidden)]
pub fn geometry_row_storage_allocations() -> u64 {
    ROW_STORAGE_ALLOCATIONS.get()
}

/// default column width in characters of max-digit-width (excel default).
pub const DEFAULT_COL_WIDTH_CHARS: f64 = 8.43;
/// default row height in points (excel default for 11pt calibri).
pub const DEFAULT_ROW_HEIGHT_PT: f64 = 15.0;
/// point size assumed for a cell whose style names none.
const DEFAULT_FONT_SIZE_PT: f64 = 11.0;
/// max-digit-width of the default font (calibri 11) in pixels at 96dpi.
pub const MAX_DIGIT_WIDTH_PX: f64 = 7.0;
/// css/screen pixels per point at 96dpi (96/72).
pub const PX_PER_PT: f64 = 96.0 / 72.0;
pub const EMU_PER_PT: f64 = 12_700.0;

/// column width in characters -> pixels per ecma-376 §18.3.1.13: the stored
/// width folds in 5px of padding; reverse it and snap to excel's 1/256 grid.
pub fn col_chars_to_px(chars: f64) -> f32 {
    if chars == 0.0 {
        return 0.0;
    }
    let mdw = MAX_DIGIT_WIDTH_PX;
    (((chars * mdw + 5.0) / mdw * 256.0).round() / 256.0 * mdw) as f32
}

/// row height in points -> pixels at 96dpi.
pub fn row_pt_to_px(pt: f64) -> f32 {
    (pt * PX_PER_PT) as f32
}

/// height in points excel gives a row auto-fitted to `size_pt` text, on
/// excel's quarter-point grid. the constants are fitted to the row bands in
/// excel's own renders of the fidelity corpus.
pub fn autofit_row_height_pt(size_pt: f64) -> f64 {
    (((size_pt * AUTOFIT_LINE_RATIO - AUTOFIT_LEADING_PT) * 4.0).round() / 4.0).max(0.0)
}

const AUTOFIT_LINE_RATIO: f64 = 1.4615;
const AUTOFIT_LEADING_PT: f64 = 2.0769;

/// floors to whole points, absorbing the float error a ratio leaves just under
/// an exact integer (90.0 x 14/15 evaluates to 62.999...).
fn floor_pt(pt: f64) -> f64 {
    (pt + 1e-9).floor()
}

/// the font `cellXfs[0]` resolves to: what a cell inheriting the normal style renders in.
fn normal_font(styles: &Stylesheet) -> Option<&Font> {
    styles.font_for(0).or_else(|| styles.fonts.first())
}

/// the normal font the row grid measures with; on the print path it is the
/// same value the text falls back to, so the two cannot disagree.
#[derive(Debug, Clone, Copy)]
struct NormalFace<'a> {
    size_pt: f64,
    family: Option<&'a str>,
}

impl<'a> NormalFace<'a> {
    fn from_styles(styles: &'a Stylesheet) -> Self {
        let font = normal_font(styles);
        Self {
            size_pt: font
                .and_then(|font| font.size_pt)
                .filter(|pt| pt.is_finite() && *pt > 0.0)
                .unwrap_or(DEFAULT_FONT_SIZE_PT),
            family: font.and_then(|font| font.name.as_deref()),
        }
    }

    fn from_metrics(metrics: &'a PrintMetrics) -> Self {
        Self {
            size_pt: Some(metrics.font_size_pt as f64)
                .filter(|pt| pt.is_finite() && *pt > 0.0)
                .unwrap_or(DEFAULT_FONT_SIZE_PT),
            family: Some(metrics.font_family.as_str()),
        }
    }
}

/// the factor excel applies to every stored `ht` when `defaultRowHeight` is a
/// cached hint it recomputes; `None` keeps stored heights. gated on calibri
/// because [`autofit_row_height_pt`] is a calibri measurement.
fn stored_height_scale(sheet: &Sheet, normal: NormalFace<'_>) -> Option<f64> {
    if sheet.format.custom_height {
        return None;
    }
    let declared = sheet
        .format
        .default_row_height_pt
        .filter(|pt| pt.is_finite() && *pt > 0.0)?;
    if !normal
        .family
        .is_none_or(|name| name.eq_ignore_ascii_case("calibri"))
    {
        return None;
    }
    let fitted = autofit_row_height_pt(normal.size_pt);
    (fitted.is_finite() && fitted > 0.0 && fitted != declared).then_some(fitted / declared)
}

/// the height a cell's font contributes to row autofit, if any: a declared
/// `defaultRowHeight` makes every styled cell count, otherwise only fonts
/// taller than the sheet default grow a row.
fn autofit_height(
    styles: &Stylesheet,
    style: Option<u32>,
    default_pt: f64,
    cached_default: bool,
    normal: NormalFace<'_>,
) -> Option<f64> {
    let size = style
        .and_then(|style| styles.font_for(style))
        .and_then(|font| font.size_pt)
        .filter(|pt| pt.is_finite() && *pt > 0.0)
        .unwrap_or(normal.size_pt);
    let height = autofit_row_height_pt(size);
    (cached_default || height > default_pt).then_some(height)
}

/// whether a cell at `at` carrying `style` participates in row autofit —
/// the single-cell form of the skip rules in `autofit_rows`; keep in sync.
pub fn autofit_relevant(
    sheet: &Sheet,
    styles: &Stylesheet,
    at: CellRef,
    style: Option<u32>,
) -> bool {
    if sheet.format.custom_height
        || sheet.row_heights.contains_key(&at.row)
        || sheet
            .merges
            .iter()
            .any(|range| range.end.row > range.start.row && range.start == at)
    {
        return false;
    }
    let default_pt = sheet
        .format
        .default_row_height_pt
        .unwrap_or(DEFAULT_ROW_HEIGHT_PT);
    autofit_height(
        styles,
        style,
        default_pt,
        sheet.format.default_row_height_pt.is_some(),
        NormalFace::from_styles(styles),
    )
    .is_some()
}

/// rows excel auto-fits, with the height each takes: every row that carries no
/// `ht` and is not pinned by `sheetFormatPr/@customHeight`. a declared
/// `defaultRowHeight` is a cached hint excel recomputes, so a row with content
/// takes its fitted height outright; without one the sheet default is already
/// the normal font's fitted height and only taller content grows a row.
fn autofit_rows(
    sheet: &Sheet,
    styles: &Stylesheet,
    default_pt: f64,
    normal: NormalFace<'_>,
) -> BTreeMap<RowId, f64> {
    let mut fitted = BTreeMap::new();
    if sheet.format.custom_height {
        return fitted;
    }
    let cached_default = sheet.format.default_row_height_pt.is_some();
    let spanned: HashSet<(RowId, ColId)> = sheet
        .merges
        .iter()
        .filter(|range| range.end.row > range.start.row)
        .map(|range| (range.start.row, range.start.col))
        .collect();
    for (at, cell) in sheet.iter_cells() {
        #[cfg(feature = "test-counters")]
        AUTOFIT_CELL_VISITS.set(AUTOFIT_CELL_VISITS.get().wrapping_add(1));
        if sheet.row_heights.contains_key(&at.row) || spanned.contains(&(at.row, at.col)) {
            continue;
        }
        // a cell naming no style still inherits its column's
        let style = cell.style.or_else(|| sheet.col_style(at.col));
        if let Some(height) = autofit_height(styles, style, default_pt, cached_default, normal) {
            let entry = fitted.entry(at.row).or_insert(height);
            if height > *entry {
                *entry = height;
            }
        }
    }
    fitted
}

/// EMU -> pixels in the renderer's 96dpi coordinate space.
pub fn emu_to_px(emu: i64) -> f64 {
    emu as f64 / EMU_PER_PT * PX_PER_PT
}

#[derive(Debug, Clone, serde::Deserialize, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PrintMetrics {
    pub dpi: f32,
    pub max_digit_width: f32,
    pub default_row_height_pt: f32,
    pub default_column_width: Option<f64>,
    pub font_size_pt: f32,
    pub font_family: String,
    pub font_ascent: f32,
    pub font_descent: f32,
}

impl PrintMetrics {
    /// a zero default row height is how a sheet declares `zeroHeight="1"`:
    /// every row without an explicit `ht` collapses, exactly as a zero default
    /// column width already collapses unsized columns.
    pub fn is_valid(&self) -> bool {
        !self.font_family.is_empty()
            && self.font_family.len() <= 1024
            && (36.0..=600.0).contains(&self.dpi)
            && (1.0..=256.0).contains(&self.max_digit_width)
            && (0.0..=409.0).contains(&self.default_row_height_pt)
            && self
                .default_column_width
                .is_none_or(|w| (0.0..=255.0).contains(&w))
            && (1.0..=409.0).contains(&self.font_size_pt)
            && (0.0..=4096.0).contains(&self.font_ascent)
            && (0.0..=4096.0).contains(&self.font_descent)
            && self.font_ascent + self.font_descent > 0.0
    }

    pub fn column_pixels(&self, width: f64) -> f32 {
        let digit = self.max_digit_width as f64;
        (((256.0 * width + (128.0 / digit).floor()) / 256.0 * digit).floor() * 96.0
            / self.dpi as f64) as f32
    }
}

const SIZE_CHUNK_LEN: usize = 256;

#[derive(Debug, Clone)]
struct SizeChunk {
    sum: f64,
    count: u32,
    values: [f32; SIZE_CHUNK_LEN],
}

impl Default for SizeChunk {
    fn default() -> Self {
        Self {
            sum: 0.0,
            count: 0,
            values: [f32::NAN; SIZE_CHUNK_LEN],
        }
    }
}

#[derive(Debug)]
struct Sizes {
    chunks: Vec<SizeChunk>,
    limit: u32,
    default: f32,
    end: u32,
    count: u32,
}

impl Clone for Sizes {
    fn clone(&self) -> Self {
        #[cfg(feature = "test-counters")]
        if self.limit == xlsx_model::MAX_ROWS && !self.chunks.is_empty() {
            ROW_STORAGE_ALLOCATIONS.set(ROW_STORAGE_ALLOCATIONS.get().wrapping_add(1));
        }
        Self {
            chunks: self.chunks.clone(),
            limit: self.limit,
            default: self.default,
            end: self.end,
            count: self.count,
        }
    }
}

impl Sizes {
    fn from_values(limit: u32, default: f32, values: &[(u32, f32)]) -> Self {
        let end = values.last().map_or(0, |(index, _)| index + 1);
        #[cfg(feature = "test-counters")]
        if limit == xlsx_model::MAX_ROWS && end > 0 {
            ROW_STORAGE_ALLOCATIONS.set(ROW_STORAGE_ALLOCATIONS.get().wrapping_add(1));
        }
        let mut chunks = vec![SizeChunk::default(); (end as usize).div_ceil(SIZE_CHUNK_LEN)];
        for &(index, value) in values {
            let chunk = &mut chunks[index as usize / SIZE_CHUNK_LEN];
            chunk.values[index as usize % SIZE_CHUNK_LEN] = value;
            chunk.sum += f64::from(value);
            chunk.count += 1;
        }
        for index in 1..=chunks.len() {
            let parent = index + index.isolate_lowest_one();
            if parent <= chunks.len() {
                chunks[parent - 1].sum += chunks[index - 1].sum;
                chunks[parent - 1].count += chunks[index - 1].count;
            }
        }
        Self {
            chunks,
            limit,
            default,
            end,
            count: values.len() as u32,
        }
    }

    fn prefix(&self, count: usize) -> (f64, u32) {
        let mut index = count.min(self.chunks.len());
        let mut sum = 0.0;
        let mut count = 0;
        while index > 0 {
            #[cfg(feature = "test-counters")]
            PREFIX_QUERIES.set(PREFIX_QUERIES.get().wrapping_add(1));
            sum += self.chunks[index - 1].sum;
            count += self.chunks[index - 1].count;
            index &= index - 1;
        }
        (sum, count)
    }

    fn grow(&mut self, count: usize) {
        if count <= self.chunks.len() {
            return;
        }
        let old_len = self.chunks.len();
        let (sum, active) = self.prefix(old_len);
        let count = count
            .next_power_of_two()
            .min((self.limit as usize).div_ceil(SIZE_CHUNK_LEN));
        #[cfg(feature = "test-counters")]
        if self.limit == xlsx_model::MAX_ROWS && count > self.chunks.capacity() {
            ROW_STORAGE_ALLOCATIONS.set(ROW_STORAGE_ALLOCATIONS.get().wrapping_add(1));
        }
        self.chunks.resize(count, SizeChunk::default());
        for index in old_len..count {
            let start = index & (index + 1);
            if start < old_len {
                let (before, preceding) = self.prefix(start);
                self.chunks[index].sum = sum - before;
                self.chunks[index].count = active - preceding;
            }
        }
    }

    fn set(&mut self, index: u32, value: Option<f32>) {
        assert!(index < self.limit);
        let old = self.value(index);
        let chunk = index as usize / SIZE_CHUNK_LEN;
        self.grow(chunk + 1);
        self.chunks[chunk].values[index as usize % SIZE_CHUNK_LEN] = value.unwrap_or(f32::NAN);
        let delta = f64::from(value.unwrap_or(0.0)) - f64::from(old.unwrap_or(0.0));
        let added = u32::from(value.is_some() && old.is_none());
        let removed = u32::from(value.is_none() && old.is_some());
        let mut current = chunk + 1;
        while current <= self.chunks.len() {
            #[cfg(feature = "test-counters")]
            PREFIX_UPDATES.set(PREFIX_UPDATES.get().wrapping_add(1));
            self.chunks[current - 1].sum += delta;
            self.chunks[current - 1].count += added;
            self.chunks[current - 1].count -= removed;
            current += current.isolate_lowest_one();
        }
        self.count += added;
        self.count -= removed;
        if value.is_some() {
            self.end = self.end.max(index + 1);
        } else if index + 1 == self.end {
            self.end = self.last();
        }
    }

    fn last(&self) -> u32 {
        if self.count == 0 {
            return 0;
        }
        let mut index = 0;
        let mut count = 0;
        let mut step = 1 << self.chunks.len().ilog2();
        while step > 0 {
            #[cfg(feature = "test-counters")]
            PREFIX_QUERIES.set(PREFIX_QUERIES.get().wrapping_add(1));
            let next = index + step;
            if next <= self.chunks.len() && count + self.chunks[next - 1].count < self.count {
                count += self.chunks[next - 1].count;
                index = next;
            }
            step >>= 1;
        }
        let last = self.chunks[index]
            .values
            .iter()
            .rposition(|value| !value.is_nan())
            .unwrap();
        (index * SIZE_CHUNK_LEN + last + 1) as u32
    }

    fn value(&self, index: u32) -> Option<f32> {
        let chunk = self.chunks.get(index as usize / SIZE_CHUNK_LEN)?;
        let value = chunk.values[index as usize % SIZE_CHUNK_LEN];
        (!value.is_nan()).then_some(value)
    }

    fn offset(&self, index: u32) -> f32 {
        let chunk = index as usize / SIZE_CHUNK_LEN;
        let (mut sum, mut count) = self.prefix(chunk);
        if let Some(chunk) = self.chunks.get(chunk) {
            #[cfg(feature = "test-counters")]
            PREFIX_QUERIES.set(PREFIX_QUERIES.get().wrapping_add(1));
            for &value in &chunk.values[..index as usize % SIZE_CHUNK_LEN] {
                if !value.is_nan() {
                    sum += f64::from(value);
                    count += 1;
                }
            }
        }
        self.boundary(index, sum, count)
    }

    fn boundary(&self, index: u32, sum: f64, count: u32) -> f32 {
        if index == count {
            sum as f32
        } else {
            (sum + f64::from(self.default) * f64::from(index - count)) as f32
        }
    }

    fn at(&self, position: f32) -> u32 {
        if position.is_nan() || position < 0.0 {
            return 0;
        }
        let edge = self.offset(self.end);
        if position >= edge {
            if self.default == 0.0 {
                return self.end;
            }
            let (sum, count) = self.prefix(self.chunks.len());
            let mut index = self
                .end
                .saturating_add(((position - edge) / self.default) as u32);
            while index > self.end && position < self.boundary(index, sum, count) {
                index -= 1;
            }
            while index < u32::MAX && position >= self.boundary(index + 1, sum, count) {
                index += 1;
            }
            return index;
        }
        let mut index = 0;
        let mut sum = 0.0;
        let mut count = 0;
        let mut step = 1 << self.chunks.len().ilog2();
        while step > 0 {
            #[cfg(feature = "test-counters")]
            PREFIX_QUERIES.set(PREFIX_QUERIES.get().wrapping_add(1));
            let next = index + step;
            if next <= self.chunks.len() {
                let next_sum = sum + self.chunks[next - 1].sum;
                let next_count = count + self.chunks[next - 1].count;
                let boundary = self.boundary((next * SIZE_CHUNK_LEN) as u32, next_sum, next_count);
                if position >= boundary {
                    sum = next_sum;
                    count = next_count;
                    index = next;
                }
            }
            step >>= 1;
        }
        let mut offset = sum;
        let implicit = (index * SIZE_CHUNK_LEN) as u32 - count;
        if implicit > 0 {
            offset += f64::from(self.default) * f64::from(implicit);
        }
        for (within, &value) in self.chunks[index].values.iter().enumerate() {
            offset += f64::from(if value.is_nan() { self.default } else { value });
            if position < offset as f32 {
                return (index * SIZE_CHUNK_LEN + within) as u32;
            }
        }
        ((index + 1) * SIZE_CHUNK_LEN) as u32
    }
}

#[derive(Debug, Clone)]
pub struct GridGeometry {
    columns: Sizes,
    rows: Sizes,
    spanned: Arc<HashSet<(RowId, ColId)>>,
}

impl GridGeometry {
    /// build cumulative offset tables from a sheet's custom widths/heights.
    pub fn new(sheet: &Sheet, styles: &Stylesheet) -> Self {
        Self::with_sizes(
            sheet,
            styles,
            col_chars_to_px(DEFAULT_COL_WIDTH_CHARS),
            sheet
                .format
                .default_row_height_pt
                .unwrap_or(DEFAULT_ROW_HEIGHT_PT),
            NormalFace::from_styles(styles),
            col_chars_to_px,
        )
    }

    pub fn for_print(sheet: &Sheet, styles: &Stylesheet, metrics: &PrintMetrics) -> Self {
        Self::with_sizes(
            sheet,
            styles,
            metrics
                .default_column_width
                .map_or(col_chars_to_px(DEFAULT_COL_WIDTH_CHARS), |width| {
                    metrics.column_pixels(width)
                }),
            metrics.default_row_height_pt as f64,
            NormalFace::from_metrics(metrics),
            |width| metrics.column_pixels(width),
        )
    }

    fn with_sizes(
        sheet: &Sheet,
        styles: &Stylesheet,
        default_col_px: f32,
        default_row_pt: f64,
        normal: NormalFace<'_>,
        column_pixels: impl Fn(f64) -> f32,
    ) -> Self {
        #[cfg(feature = "test-counters")]
        GEOMETRY_CONSTRUCTIONS.set(GEOMETRY_CONSTRUCTIONS.get().wrapping_add(1));
        let default_row_px = if sheet.format.zero_height {
            0.0
        } else {
            row_pt_to_px(default_row_pt)
        };
        let fitted = autofit_rows(sheet, styles, default_row_pt, normal);
        let scale = stored_height_scale(sheet, normal);
        let columns: Vec<_> = sheet
            .col_widths
            .iter()
            .map(|(&col, &width)| {
                (
                    col,
                    if sheet.col_hidden(col) {
                        0.0
                    } else {
                        column_pixels(width)
                    },
                )
            })
            .collect();
        let sized: std::collections::BTreeSet<_> = sheet
            .row_heights
            .keys()
            .chain(fitted.keys())
            .copied()
            .collect();
        let rows: Vec<_> = sized
            .into_iter()
            .filter_map(|row| {
                row_size(sheet, row, fitted.get(&row).copied(), scale).map(|size| (row, size))
            })
            .collect();
        let columns = Sizes::from_values(xlsx_model::MAX_COLS, default_col_px, &columns);
        let rows = Sizes::from_values(xlsx_model::MAX_ROWS, default_row_px, &rows);
        let spanned = Arc::new(
            sheet
                .merges
                .iter()
                .filter(|range| range.end.row > range.start.row)
                .map(|range| (range.start.row, range.start.col))
                .collect(),
        );
        Self {
            columns,
            rows,
            spanned,
        }
    }

    #[doc(hidden)]
    pub fn update_merges(&mut self, sheet: &Sheet) {
        self.spanned = Arc::new(
            sheet
                .merges
                .iter()
                .filter(|range| range.end.row > range.start.row)
                .map(|range| (range.start.row, range.start.col))
                .collect(),
        );
    }

    #[doc(hidden)]
    pub fn update_rows(
        &mut self,
        sheet: &Sheet,
        styles: &Stylesheet,
        rows: impl IntoIterator<Item = RowId>,
    ) {
        let normal = NormalFace::from_styles(styles);
        let default_pt = sheet
            .format
            .default_row_height_pt
            .unwrap_or(DEFAULT_ROW_HEIGHT_PT);
        let scale = stored_height_scale(sheet, normal);
        for row in rows {
            let mut fitted: Option<f64> = None;
            if !sheet.format.custom_height && !sheet.row_heights.contains_key(&row) {
                for (at, cell) in sheet.iter_cells_in_rect(row..row + 1, 0..xlsx_model::MAX_COLS) {
                    #[cfg(feature = "test-counters")]
                    AUTOFIT_CELL_VISITS.set(AUTOFIT_CELL_VISITS.get().wrapping_add(1));
                    if self.spanned.contains(&(at.row, at.col)) {
                        continue;
                    }
                    let style = cell.style.or_else(|| sheet.col_style(at.col));
                    if let Some(height) = autofit_height(
                        styles,
                        style,
                        default_pt,
                        sheet.format.default_row_height_pt.is_some(),
                        normal,
                    ) {
                        fitted = Some(fitted.map_or(height, |old| old.max(height)));
                    }
                }
            }
            let size = row_size(sheet, row, fitted, scale);
            if self.rows.value(row) != size {
                self.rows.set(row, size);
            }
        }
    }

    #[doc(hidden)]
    pub fn update_columns(&mut self, sheet: &Sheet, columns: impl IntoIterator<Item = ColId>) {
        for col in columns {
            let size = sheet.col_widths.get(&col).map(|&width| {
                if sheet.col_hidden(col) {
                    0.0
                } else {
                    col_chars_to_px(width)
                }
            });
            if self.columns.value(col) != size {
                self.columns.set(col, size);
            }
        }
    }

    pub fn col_x(&self, col: ColId) -> f32 {
        self.columns.offset(col)
    }

    pub fn row_y(&self, row: RowId) -> f32 {
        self.rows.offset(row)
    }

    pub fn col_at_x(&self, x: f32) -> ColId {
        self.columns.at(x)
    }

    pub fn row_at_y(&self, y: f32) -> RowId {
        self.rows.at(y)
    }

    /// half-open (row, col) ranges of cells intersecting the viewport.
    pub fn viewport_range(&self, vp: &Viewport) -> (Range<RowId>, Range<ColId>) {
        let bottom = self.row_y(MAX_ROWS);
        let right = self.col_x(MAX_COLS);
        let rows = if bottom - vp.y <= 0.0 {
            MAX_ROWS..MAX_ROWS
        } else {
            let r0 = self.row_at_y(vp.y).min(MAX_ROWS - 1);
            let r1 = self
                .row_at_y((vp.y + vp.height).min(bottom))
                .saturating_add(1)
                .min(MAX_ROWS);
            r0..r1
        };
        let cols = if right - vp.x <= 0.0 {
            MAX_COLS..MAX_COLS
        } else {
            let c0 = self.col_at_x(vp.x).min(MAX_COLS - 1);
            let c1 = self
                .col_at_x((vp.x + vp.width).min(right))
                .saturating_add(1)
                .min(MAX_COLS);
            c0..c1
        };
        (rows, cols)
    }
}

fn row_size(sheet: &Sheet, row: RowId, fitted: Option<f64>, scale: Option<f64>) -> Option<f32> {
    sheet
        .row_heights
        .get(&row)
        .copied()
        .map(|height| scale.map_or(height, |scale| floor_pt(height * scale)))
        .or(fitted)
        .map(|height| {
            if sheet.row_hidden(row) {
                0.0
            } else {
                row_pt_to_px(height)
            }
        })
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::BTreeMap;
    use xlsx_model::CellRef;
    use xlsx_model::styles::{Font, Xf};

    fn styles() -> Stylesheet {
        Stylesheet::default()
    }

    fn sheet_with(cols: &[(ColId, f64)], rows: &[(RowId, f64)]) -> Sheet {
        let mut s = Sheet::new("S");
        s.col_widths = cols.iter().copied().collect::<BTreeMap<_, _>>();
        s.row_heights = rows.iter().copied().collect::<BTreeMap<_, _>>();
        s
    }

    #[test]
    fn inverse_lookup_matches_fractional_offsets_in_default_tails() {
        let sheet = sheet_with(&[(0, 13.25)], &[(0, 25.25)]);
        let grid = GridGeometry::new(&sheet, &styles());
        for row in [1, 10, 1_000, 100_000, xlsx_model::MAX_ROWS - 1] {
            assert_eq!(grid.row_at_y(grid.row_y(row)), row);
        }
        for col in 1..xlsx_model::MAX_COLS {
            assert_eq!(grid.col_at_x(grid.col_x(col)), col);
        }
    }

    #[test]
    fn fractional_defaults_round_trip_published_boundaries() {
        for default_pt in [15.25, 15.5, 12.75, 17.1] {
            let mut sheet = sheet_with(&[(0, 13.25)], &[(0, 25.25)]);
            sheet.format.custom_height = true;
            sheet.format.default_row_height_pt = Some(default_pt);
            let grid = GridGeometry::new(&sheet, &styles());
            for row in (0..1_024).chain([100_000, xlsx_model::MAX_ROWS - 1]) {
                assert_eq!(grid.row_at_y(grid.row_y(row)), row, "default {default_pt}");
            }
            let grid = GridGeometry::with_sizes(
                &sheet,
                &styles(),
                row_pt_to_px(default_pt),
                default_pt,
                NormalFace::from_styles(&styles()),
                col_chars_to_px,
            );
            for col in 0..xlsx_model::MAX_COLS {
                assert_eq!(grid.col_at_x(grid.col_x(col)), col, "default {default_pt}");
            }
        }
    }

    #[test]
    fn dense_sizes_use_one_compact_allocation() {
        for rows in [100_000, xlsx_model::MAX_ROWS] {
            let values: Vec<_> = (0..rows).map(|row| (row, 18.0)).collect();
            let sizes = Sizes::from_values(xlsx_model::MAX_ROWS, 20.0, &values);
            let bytes = sizes.chunks.capacity() * std::mem::size_of::<SizeChunk>();
            assert!(bytes <= rows as usize * 5, "allocated {bytes} bytes");
            assert_eq!(sizes.offset(rows), rows as f32 * 18.0);
        }
        let sizes = Sizes::from_values(xlsx_model::MAX_ROWS, 20.0, &[]);
        assert_eq!(sizes.chunks.capacity(), 0);
    }

    #[test]
    fn size_updates_preserve_chunk_sums_and_explicit_bounds() {
        let mut sizes = Sizes::from_values(xlsx_model::MAX_ROWS, 0.0, &[(0, 3.0)]);
        for row in [255, 256, 100_000, xlsx_model::MAX_ROWS - 1] {
            sizes.set(row, Some(7.0));
            assert_eq!(sizes.end, row + 1);
            assert_eq!(sizes.at(sizes.offset(row)), row);
        }
        assert_eq!(sizes.offset(xlsx_model::MAX_ROWS), 31.0);
        sizes.set(256, Some(0.0));
        assert_eq!(sizes.offset(xlsx_model::MAX_ROWS), 24.0);
        for (row, end) in [
            (xlsx_model::MAX_ROWS - 1, 100_001),
            (100_000, 257),
            (256, 256),
            (255, 1),
            (0, 0),
        ] {
            sizes.set(row, None);
            assert_eq!(sizes.end, end);
            assert_eq!(sizes.at(sizes.offset(end)), end);
        }
        assert_eq!(sizes.offset(xlsx_model::MAX_ROWS), 0.0);
    }

    #[test]
    fn autofit_matches_the_row_bands_excel_renders() {
        assert_eq!(autofit_row_height_pt(11.0), 14.0);
        assert_eq!(autofit_row_height_pt(12.0), 15.5);
        assert_eq!(autofit_row_height_pt(24.0), 33.0);
        assert_eq!(autofit_row_height_pt(30.0), 41.75);
        assert_eq!(autofit_row_height_pt(0.0), 0.0);
    }

    #[test]
    fn an_unsized_row_takes_its_tallest_font_unless_the_sheet_pins_it() {
        let mut styles = Stylesheet::default();
        styles.fonts.push(Font {
            size_pt: Some(30.0),
            ..Font::default()
        });
        styles.cell_xfs.push(Xf {
            font: Some(0),
            ..Xf::default()
        });
        let mut sheet = Sheet::new("S");
        sheet.set_cell(
            CellRef::new(1, 0),
            xlsx_model::workbook::Cell {
                style: Some(0),
                ..Default::default()
            },
        );
        let default_px = row_pt_to_px(DEFAULT_ROW_HEIGHT_PT);
        let grown = GridGeometry::new(&sheet, &styles);
        assert_eq!(grown.row_y(1), default_px);
        assert!(
            (grown.row_y(2) - grown.row_y(1) - row_pt_to_px(autofit_row_height_pt(30.0))).abs()
                < 0.001
        );

        sheet.format.custom_height = true;
        let pinned = GridGeometry::new(&sheet, &styles);
        assert_eq!(pinned.row_y(2), default_px * 2.0);

        sheet.format.default_row_height_pt = Some(DEFAULT_ROW_HEIGHT_PT);
        let cached = GridGeometry::new(&sheet, &styles);
        assert_eq!(cached.row_y(2), default_px * 2.0);

        sheet.format.custom_height = false;
        let recomputed = GridGeometry::new(&sheet, &styles);
        assert!(
            (recomputed.row_y(2) - recomputed.row_y(1) - row_pt_to_px(autofit_row_height_pt(30.0)))
                .abs()
                < 0.001
        );

        sheet.format.default_row_height_pt = None;
        sheet.row_heights.insert(1, 9.0);
        let authored = GridGeometry::new(&sheet, &styles);
        assert!((authored.row_y(2) - authored.row_y(1) - row_pt_to_px(9.0)).abs() < 0.001);
    }

    #[test]
    fn stored_heights_rescale_with_a_recomputed_cached_default() {
        let mut styles = Stylesheet::default();
        styles.fonts.push(Font {
            name: Some("Calibri".into()),
            size_pt: Some(11.0),
            ..Font::default()
        });
        let mut sheet = sheet_with(&[], &[(0, 15.0), (1, 24.0), (2, 37.5)]);
        sheet.format.default_row_height_pt = Some(15.0);
        let rescaled = GridGeometry::new(&sheet, &styles);
        for (row, pt) in [(0, 14.0), (1, 22.0), (2, 35.0)] {
            assert!(
                (rescaled.row_y(row + 1) - rescaled.row_y(row) - row_pt_to_px(pt)).abs() < 0.001
            );
        }

        sheet.format.custom_height = true;
        let pinned = GridGeometry::new(&sheet, &styles);
        assert!((pinned.row_y(1) - row_pt_to_px(15.0)).abs() < 0.001);

        sheet.format.custom_height = false;
        styles.fonts[0].name = Some("Arial".into());
        let unmeasured = GridGeometry::new(&sheet, &styles);
        assert!((unmeasured.row_y(1) - row_pt_to_px(15.0)).abs() < 0.001);
    }

    #[test]
    fn an_unstyled_row_fits_the_normal_size_not_eleven_points() {
        let mut styles = Stylesheet::default();
        styles.fonts.push(Font {
            name: Some("Arial".into()),
            size_pt: Some(10.0),
            ..Font::default()
        });
        styles.cell_xfs.push(Xf {
            font: Some(0),
            ..Xf::default()
        });
        let mut sheet = Sheet::new("S");
        sheet.format.default_row_height_pt = Some(12.75);
        sheet.set_cell(
            CellRef::new(0, 0),
            xlsx_model::workbook::Cell {
                style: Some(0),
                ..Default::default()
            },
        );
        let grid = GridGeometry::new(&sheet, &styles);
        assert!((grid.row_y(1) - row_pt_to_px(autofit_row_height_pt(10.0))).abs() < 0.001);
        assert!((grid.row_y(1) - row_pt_to_px(autofit_row_height_pt(11.0))).abs() > 1.0);
    }

    #[test]
    fn printing_ignores_a_non_finite_metric_font_size() {
        let metrics = PrintMetrics {
            dpi: 72.0,
            max_digit_width: 6.0,
            default_row_height_pt: 12.75,
            default_column_width: None,
            font_size_pt: f32::INFINITY,
            font_family: "Calibri".into(),
            font_ascent: 14.0,
            font_descent: 4.0,
        };
        let mut sheet = Sheet::new("S");
        sheet.format.default_row_height_pt = Some(12.75);
        sheet.set_cell(
            CellRef::new(0, 0),
            xlsx_model::workbook::Cell {
                value: xlsx_model::value::CellValue::Number { value: 1.0 },
                ..Default::default()
            },
        );
        let grid = GridGeometry::for_print(&sheet, &Stylesheet::default(), &metrics);
        assert!(grid.row_y(1).is_finite());
        assert!(
            (grid.row_y(1) - row_pt_to_px(autofit_row_height_pt(DEFAULT_FONT_SIZE_PT))).abs()
                < 0.001
        );
    }

    #[test]
    fn printing_fits_an_unstyled_row_to_the_metrics_the_text_falls_back_to() {
        let metrics = PrintMetrics {
            dpi: 72.0,
            max_digit_width: 6.0,
            default_row_height_pt: 12.75,
            default_column_width: None,
            font_size_pt: 16.0,
            font_family: "Arial".into(),
            font_ascent: 14.0,
            font_descent: 4.0,
        };
        let mut sheet = Sheet::new("S");
        sheet.format.default_row_height_pt = Some(12.75);
        sheet.set_cell(
            CellRef::new(0, 0),
            xlsx_model::workbook::Cell {
                value: xlsx_model::value::CellValue::Number { value: 1.0 },
                ..Default::default()
            },
        );
        let grid = GridGeometry::for_print(&sheet, &Stylesheet::default(), &metrics);
        assert!((grid.row_y(1) - row_pt_to_px(autofit_row_height_pt(16.0))).abs() < 0.001);
    }

    #[test]
    fn the_normal_face_comes_from_the_first_cell_format_not_the_font_table() {
        let mut styles = Stylesheet::default();
        styles.fonts.push(Font {
            name: Some("Calibri".into()),
            size_pt: Some(11.0),
            ..Font::default()
        });
        styles.fonts.push(Font {
            name: Some("Comic Sans MS".into()),
            size_pt: Some(30.0),
            ..Font::default()
        });
        styles.cell_xfs.push(Xf {
            font: Some(1),
            ..Xf::default()
        });
        let mut sheet = sheet_with(&[], &[(0, 15.0)]);
        sheet.format.default_row_height_pt = Some(15.0);
        let grid = GridGeometry::new(&sheet, &styles);
        assert!((grid.row_y(1) - row_pt_to_px(15.0)).abs() < 0.001);
    }

    #[test]
    fn a_non_finite_measurement_leaves_stored_heights_alone() {
        let mut styles = Stylesheet::default();
        styles.fonts.push(Font {
            name: Some("Calibri".into()),
            size_pt: Some(f64::INFINITY),
            ..Font::default()
        });
        let mut sheet = sheet_with(&[], &[(0, 15.0)]);
        sheet.format.default_row_height_pt = Some(15.0);
        assert!((GridGeometry::new(&sheet, &styles).row_y(1) - row_pt_to_px(14.0)).abs() < 0.001);

        sheet.format.default_row_height_pt = Some(f64::INFINITY);
        styles.fonts[0].size_pt = Some(11.0);
        assert!((GridGeometry::new(&sheet, &styles).row_y(1) - row_pt_to_px(15.0)).abs() < 0.001);
    }

    #[test]
    fn a_rescaled_height_landing_on_an_integer_keeps_it() {
        let mut styles = Stylesheet::default();
        styles.fonts.push(Font {
            name: Some("Calibri".into()),
            size_pt: Some(11.0),
            ..Font::default()
        });
        let mut sheet = sheet_with(&[], &[(0, 90.0)]);
        sheet.format.default_row_height_pt = Some(20.0);
        let grid = GridGeometry::new(&sheet, &styles);
        assert!((grid.row_y(1) - row_pt_to_px(63.0)).abs() < 0.001);
    }

    /// `zeroHeight` hides rows without their own `ht`; a sized row still shows.
    #[test]
    fn zero_height_collapses_only_the_unsized_rows() {
        let mut styles = Stylesheet::default();
        styles.fonts.push(Font {
            size_pt: Some(30.0),
            ..Font::default()
        });
        styles.cell_xfs.push(Xf {
            font: Some(0),
            ..Xf::default()
        });
        let mut sheet = sheet_with(&[], &[(1, 24.0)]);
        sheet.format.zero_height = true;
        sheet.set_cell(
            CellRef::new(0, 0),
            xlsx_model::workbook::Cell {
                style: Some(0),
                ..Default::default()
            },
        );
        let grid = GridGeometry::new(&sheet, &styles);
        assert_eq!(grid.row_y(1), 0.0);
        assert!((grid.row_y(2) - grid.row_y(1) - row_pt_to_px(24.0)).abs() < 0.001);
    }

    #[test]
    fn unit_conversion_anchors() {
        let px = col_chars_to_px(DEFAULT_COL_WIDTH_CHARS);
        assert!((px - 64.0).abs() < 0.1, "default col width was {px}");
        assert_eq!(row_pt_to_px(DEFAULT_ROW_HEIGHT_PT), 20.0);
        assert_eq!(col_chars_to_px(0.0), 0.0);
        assert_eq!(emu_to_px(9_525), 1.0);
    }

    #[test]
    fn all_default_grid_extrapolates() {
        let g = GridGeometry::new(&Sheet::new("S"), &styles());
        let dc = col_chars_to_px(DEFAULT_COL_WIDTH_CHARS);
        let dr = row_pt_to_px(DEFAULT_ROW_HEIGHT_PT);
        assert_eq!(g.col_x(0), 0.0);
        assert_eq!(g.col_x(3), 3.0 * dc);
        assert_eq!(g.row_y(10), 10.0 * dr);
        assert_eq!(g.col_at_x(dc * 2.5), 2);
        assert_eq!(g.row_at_y(dr * 4.0), 4);
        assert_eq!(g.col_at_x(-5.0), 0);
    }

    #[test]
    fn custom_widths_and_binary_search_edges() {
        let g = GridGeometry::new(&sheet_with(&[(0, 20.0), (2, 4.0)], &[(0, 30.0)]), &styles());
        let w0 = col_chars_to_px(20.0);
        let w1 = col_chars_to_px(DEFAULT_COL_WIDTH_CHARS);
        let w2 = col_chars_to_px(4.0);
        assert_eq!(g.col_x(0), 0.0);
        assert_eq!(g.col_x(1), w0);
        assert_eq!(g.col_x(2), w0 + w1);
        assert_eq!(g.col_x(3), w0 + w1 + w2);

        assert_eq!(g.col_at_x(0.0), 0);
        assert_eq!(g.col_at_x(w0), 1);
        assert_eq!(g.col_at_x(w0 + w1), 2);
        assert_eq!(g.col_at_x(w0 - 0.01), 0);
        assert_eq!(g.col_at_x(w0 + w1 + w2), 3);

        assert_eq!(g.row_y(1), row_pt_to_px(30.0));
    }

    #[test]
    fn viewport_range_covers_visible_cells() {
        let g = GridGeometry::new(&Sheet::new("S"), &styles());
        let dc = col_chars_to_px(DEFAULT_COL_WIDTH_CHARS);
        let dr = row_pt_to_px(DEFAULT_ROW_HEIGHT_PT);
        let vp = Viewport {
            x: dc * 1.5,
            y: dr * 2.5,
            width: dc * 2.0,
            height: dr * 1.0,
        };
        let (rows, cols) = g.viewport_range(&vp);
        assert_eq!(cols, 1..4);
        assert_eq!(rows, 2..4);
    }

    #[test]
    fn hidden_default_columns_preserve_visible_spans_and_end_at_the_sized_prefix() {
        let g = GridGeometry::with_sizes(
            &sheet_with(&[(1, 12.0), (3, 10.0), (4, 0.0)], &[]),
            &styles(),
            0.0,
            DEFAULT_ROW_HEIGHT_PT,
            NormalFace::from_styles(&styles()),
            col_chars_to_px,
        );
        let first_width = col_chars_to_px(12.0);
        let total_width = first_width + col_chars_to_px(10.0);
        assert_eq!(g.col_x(1), 0.0);
        assert_eq!(g.col_x(2), first_width);
        assert_eq!(g.col_x(3), first_width);
        assert_eq!(g.col_x(16_384), total_width);
        assert_eq!(g.col_at_x(-1.0), 0);
        assert_eq!(g.col_at_x(0.0), 1);
        assert_eq!(g.col_at_x(first_width), 3);
        assert_eq!(g.col_at_x(total_width - 0.01), 3);
        assert_eq!(g.col_at_x(total_width), 5);
        assert_eq!(g.col_at_x(total_width + 100.0), 5);
        assert_eq!(
            g.viewport_range(&Viewport {
                x: 0.0,
                y: 0.0,
                width: total_width + 100.0,
                height: 20.0,
            })
            .1,
            1..6
        );
    }
}
