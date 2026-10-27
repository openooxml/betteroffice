//! Sparse row and column extents.

#[cfg(feature = "test-counters")]
use std::cell::Cell;
use std::collections::{BTreeMap, HashSet};
use std::ops::Range;
use std::sync::Arc;

use xlsx_model::styles::{Font, Stylesheet};
use xlsx_model::value::CellValue;
use xlsx_model::workbook::Sheet;
use xlsx_model::{CellRef, ColId, HAlign, MAX_COLS, MAX_ROWS, RowId};

use crate::{Viewport, wrap};

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
/// excel's tallest row.
const MAX_ROW_HEIGHT_PT: f64 = 409.5;
/// `sheetFormatPr/@baseColWidth` when a sheet declares none (ecma-376 §18.3.1.81).
const DEFAULT_BASE_COL_WIDTH: f64 = 8.0;
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

/// how row heights are resolved: on screen excel's quarter-point fit, in
/// print excel's device grid at the metrics dpi, where every row lands on
/// whole device pixels and a font-sized row takes `round(1.2 x size) + 1` of
/// them (excel for mac's print rows at 72dpi: arial 10 -> 13pt, calibri 11 ->
/// 14pt, calibri or arial 12 -> 15pt).
#[derive(Debug, Clone, Copy, PartialEq)]
enum RowFit {
    Screen,
    Print { dpi: f64 },
}

impl RowFit {
    fn natural(self, size_pt: f64) -> f64 {
        match self {
            Self::Screen => autofit_row_height_pt(size_pt),
            Self::Print { dpi } => {
                (((1.2 * size_pt * dpi / 72.0).round() + 1.0) * 72.0 / dpi).max(0.0)
            }
        }
    }

    fn stored(self, height: f64, scale: Option<f64>) -> f64 {
        match self {
            Self::Screen => scale.map_or(height, |scale| floor_pt(height * scale)),
            Self::Print { dpi } => {
                (height * scale.unwrap_or(1.0) * dpi / 72.0 + 1e-9).floor() * 72.0 / dpi
            }
        }
    }

    /// the height of a row nothing sizes. in print a declared default within a
    /// point of the normal font's height is the cached fit of another
    /// platform's font metrics, which excel recomputes; one further off is an
    /// authored height it keeps, on the device grid. a sheet that declares
    /// none takes the host's `defaultRowHeightPt`.
    fn default_row(self, sheet: &Sheet, declared: f64, normal: NormalFace<'_>) -> f64 {
        let Self::Print { .. } = self else {
            return declared;
        };
        let natural = self.natural(normal.size_pt);
        if sheet.format.default_row_height_pt.is_some()
            && !sheet.format.custom_height
            && (declared - natural).abs() <= 1.0 + 1e-9
        {
            natural
        } else {
            self.stored(declared, None)
        }
    }
}

/// the baseline step of wrapped `size_pt` text: one fitted row of it, on
/// the print grid when `print_dpi` is set.
pub(crate) fn line_pitch_pt(size_pt: f64, print_dpi: Option<f64>) -> f64 {
    print_dpi
        .map_or(RowFit::Screen, |dpi| RowFit::Print { dpi })
        .natural(size_pt)
}

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

/// one `indent` level in the renderer's 96dpi space: three spaces of the
/// normal font (ecma-376 §18.8.1, whatever the cell's own font), each snapped
/// to whole device pixels as excel lays them out. excel for mac's print at
/// 72dpi indents calibri 11 by 6pt a level and arial 10 by 9pt.
pub(crate) fn indent_step_px(styles: &Stylesheet, print: Option<&PrintMetrics>) -> f32 {
    match print {
        Some(metrics) => indent_step(NormalFace::from_metrics(metrics), metrics.dpi as f64),
        None => indent_step(NormalFace::from_styles(styles), 96.0),
    }
}

/// [`indent_step_px`] of the `normal` face laid out at `dpi`.
fn indent_step(normal: NormalFace<'_>, dpi: f64) -> f32 {
    let space = (space_advance_em(normal.family) * normal.size_pt * dpi / 72.0).round();
    (3.0 * space * 96.0 / dpi) as f32
}

/// the space advance in em of common office faces and their metric twins.
fn space_advance_em(family: Option<&str>) -> f64 {
    let Some(family) = family else {
        return CALIBRI_SPACE_EM;
    };
    match family.trim().to_ascii_lowercase().as_str() {
        "calibri" | "carlito" => CALIBRI_SPACE_EM,
        "arial" | "helvetica" | "liberation sans" | "arimo" => 0.2778,
        "times new roman" | "liberation serif" | "tinos" => 0.25,
        "courier new" | "liberation mono" | "cousine" => 0.6001,
        "cambria" | "caladea" => 0.2202,
        "aptos" => 0.2031,
        "verdana" => 0.3516,
        "tahoma" => 0.3125,
        "georgia" => 0.2412,
        "trebuchet ms" => 0.3013,
        "century gothic" => 0.2769,
        "arial narrow" => 0.2280,
        "candara" => 0.2168,
        "consolas" => 0.5498,
        _ => 0.25,
    }
}

const CALIBRI_SPACE_EM: f64 = 0.2261;

/// the factor excel applies to every stored `ht` when `defaultRowHeight` is a
/// cached hint it recomputes; `None` keeps stored heights. gated on calibri
/// because [`autofit_row_height_pt`] is a calibri measurement.
fn stored_height_scale(sheet: &Sheet, normal: NormalFace<'_>, fit: RowFit) -> Option<f64> {
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
    let fitted = fit.natural(normal.size_pt);
    (fitted.is_finite() && fitted > 0.0 && fitted != declared).then_some(fitted / declared)
}

/// the height a cell's font contributes to row autofit, if any: a declared
/// `defaultRowHeight` makes every styled cell count on screen, otherwise only
/// fonts taller than the sheet default grow a row; in print only a font larger
/// than the normal one does.
fn autofit_height(
    styles: &Stylesheet,
    style: Option<u32>,
    default_pt: f64,
    cached_default: bool,
    normal: NormalFace<'_>,
    fit: RowFit,
) -> Option<f64> {
    let size = style
        .and_then(|style| styles.font_for(style))
        .and_then(|font| font.size_pt)
        .filter(|pt| pt.is_finite() && *pt > 0.0)
        .unwrap_or(normal.size_pt);
    let height = fit.natural(size);
    let grows = match fit {
        RowFit::Screen => cached_default || height > default_pt,
        RowFit::Print { .. } => size > normal.size_pt && height > default_pt,
    };
    grows.then_some(height)
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
        RowFit::Screen,
    )
    .is_some()
        || styles
            .alignment_for(style.unwrap_or(0))
            .is_some_and(wrap::wraps)
            && !sheet.merges.iter().any(|range| range.start == at)
}

/// horizontal inset of a cell's text box in px: both paddings, and one indent level.
#[derive(Debug, Clone, Copy)]
struct TextInset {
    pad: f32,
    indent: f32,
}

impl TextInset {
    /// the inset of `fit`, its indent level the step the painted text indents by.
    fn of(fit: RowFit, normal: NormalFace<'_>) -> Self {
        match fit {
            RowFit::Screen => Self {
                pad: 2.0 * crate::TEXT_PAD_PX,
                indent: indent_step(normal, 96.0),
            },
            RowFit::Print { dpi } => Self {
                pad: 5.0 * (96.0 / dpi) as f32,
                indent: indent_step(normal, dpi),
            },
        }
    }
}

/// the height a wrapped text cell asks of its row, one fitted line of its font
/// per estimated line, when it wraps onto a second line. Callers skip merge
/// anchors; a hidden column fits nothing.
fn wrapped_height(
    styles: &Stylesheet,
    style: Option<u32>,
    cell: &xlsx_model::workbook::Cell,
    column_px: f32,
    inset: TextInset,
    normal: NormalFace<'_>,
    fit: RowFit,
) -> Option<f64> {
    if !matches!(cell.value, CellValue::Text { .. }) {
        return None;
    }
    let style = style.unwrap_or(0);
    let alignment = styles
        .alignment_for(style)
        .filter(|alignment| wrap::wraps(alignment))?;
    if column_px <= 0.0 {
        return None;
    }
    let indent = match alignment.h {
        Some(HAlign::Left | HAlign::Right) => {
            alignment.indent.unwrap_or(0).min(crate::MAX_INDENT) as f32 * inset.indent
        }
        _ => 0.0,
    };
    let size = styles
        .font_for(style)
        .and_then(|font| font.size_pt)
        .filter(|pt| pt.is_finite() && *pt > 0.0)
        .unwrap_or(normal.size_pt);
    // the painter wraps the formatted text, and never beside a fill
    let formatted = xlsx_model::numfmt::format_value(
        &cell.value,
        &crate::format_code_for_style(styles, style),
        xlsx_model::DateSystem::default(),
    );
    if formatted
        .layout
        .iter()
        .any(|part| part.kind != xlsx_model::numfmt::FormatPartKind::Padding)
    {
        return None;
    }
    let lines = wrap::estimated_line_count(
        &formatted.text,
        (column_px - inset.pad - indent).max(0.0),
        size,
    );
    (lines > 1).then(|| (fit.natural(size) * lines as f64).min(MAX_ROW_HEIGHT_PT))
}

fn taller(a: Option<f64>, b: Option<f64>) -> Option<f64> {
    match (a, b) {
        (Some(a), Some(b)) => Some(a.max(b)),
        (a, b) => a.or(b),
    }
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
    column_px: impl Fn(ColId) -> f32,
    inset: TextInset,
    fit: RowFit,
) -> BTreeMap<RowId, f64> {
    let mut fitted = BTreeMap::new();
    if sheet.format.custom_height {
        return fitted;
    }
    let cached_default = fit == RowFit::Screen && sheet.format.default_row_height_pt.is_some();
    for &row in sheet.row_styles.keys() {
        if !sheet.row_heights.contains_key(&row)
            && let Some(style) = missing_row_style(sheet, row)
            && let Some(height) =
                autofit_height(styles, Some(style), default_pt, cached_default, normal, fit)
        {
            fitted.insert(row, height);
        }
    }
    let spanned = merge_anchors(sheet, true);
    let anchors = merge_anchors(sheet, false);
    for (at, cell) in sheet.iter_cells() {
        #[cfg(feature = "test-counters")]
        AUTOFIT_CELL_VISITS.set(AUTOFIT_CELL_VISITS.get().wrapping_add(1));
        if sheet.row_heights.contains_key(&at.row) || spanned.contains(&(at.row, at.col)) {
            continue;
        }
        let style = sheet.effective_style(at);
        let wrapped = wrapped_height(styles, style, cell, column_px(at.col), inset, normal, fit)
            .filter(|height| cached_default || *height > default_pt)
            .filter(|_| !anchors.contains(&(at.row, at.col)));
        if let Some(height) = taller(
            autofit_height(styles, style, default_pt, cached_default, normal, fit),
            wrapped,
        ) {
            let entry = fitted.entry(at.row).or_insert(height);
            if height > *entry {
                *entry = height;
            }
        }
    }
    fitted
}

/// the anchors of the sheet's merges, only those spanning rows when `spanning`.
fn merge_anchors(sheet: &Sheet, spanning: bool) -> HashSet<(RowId, ColId)> {
    sheet
        .merges
        .iter()
        .filter(|range| !spanning || range.end.row > range.start.row)
        .map(|range| (range.start.row, range.start.col))
        .collect()
}

fn missing_row_style(sheet: &Sheet, row: RowId) -> Option<u32> {
    if row >= MAX_ROWS || !sheet.row_styles.contains_key(&row) {
        return None;
    }
    let mut col = 0;
    for (at, _) in sheet.iter_cells_in_rect(row..row + 1, 0..xlsx_model::MAX_COLS) {
        if at.col != col {
            break;
        }
        col += 1;
    }
    (col < xlsx_model::MAX_COLS)
        .then(|| sheet.effective_style(CellRef::new(row, col)))
        .flatten()
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

    /// a column no `<col>` or `defaultColWidth` sizes: ecma-376 §18.3.1.81's
    /// `baseColWidth` digits plus 5 device pixels of padding and gridline, at
    /// the schema default `baseColWidth` of 8.
    pub fn base_column_pixels(&self) -> f32 {
        ((DEFAULT_BASE_COL_WIDTH * self.max_digit_width as f64 + 5.0) * 96.0 / self.dpi as f64)
            as f32
    }

    pub fn column_pixels(&self, width: f64) -> f32 {
        let digit = self.max_digit_width as f64;
        (((256.0 * width + (128.0 / digit).floor()) / 256.0 * digit).floor() * 96.0
            / self.dpi as f64) as f32
    }
}

const SIZE_CHUNK_LEN: usize = 256;
const _: () = assert!(
    (MAX_ROWS as usize).is_multiple_of(SIZE_CHUNK_LEN)
        && (MAX_COLS as usize).is_multiple_of(SIZE_CHUNK_LEN)
);

#[derive(Debug, Clone)]
struct SizeChunk {
    sum: f64,
    count: u32,
    /// stored values above zero.
    sized: u32,
    values: [f32; SIZE_CHUNK_LEN],
}

impl Default for SizeChunk {
    fn default() -> Self {
        Self {
            sum: 0.0,
            count: 0,
            sized: 0,
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
            chunk.sized += u32::from(value > 0.0);
        }
        for index in 1..=chunks.len() {
            let parent = index + index.isolate_lowest_one();
            if parent <= chunks.len() {
                chunks[parent - 1].sum += chunks[index - 1].sum;
                chunks[parent - 1].count += chunks[index - 1].count;
                chunks[parent - 1].sized += chunks[index - 1].sized;
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

    fn prefix(&self, count: usize) -> (f64, u32, u32) {
        let mut index = count.min(self.chunks.len());
        let mut sum = 0.0;
        let mut count = 0;
        let mut sized = 0;
        while index > 0 {
            #[cfg(feature = "test-counters")]
            PREFIX_QUERIES.set(PREFIX_QUERIES.get().wrapping_add(1));
            sum += self.chunks[index - 1].sum;
            count += self.chunks[index - 1].count;
            sized += self.chunks[index - 1].sized;
            index &= index - 1;
        }
        (sum, count, sized)
    }

    fn grow(&mut self, count: usize) {
        if count <= self.chunks.len() {
            return;
        }
        let old_len = self.chunks.len();
        let (sum, active, sized) = self.prefix(old_len);
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
                let (before, preceding, sized_before) = self.prefix(start);
                self.chunks[index].sum = sum - before;
                self.chunks[index].count = active - preceding;
                self.chunks[index].sized = sized - sized_before;
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
        let was_sized = old.is_some_and(|size| size > 0.0);
        let is_sized = value.is_some_and(|size| size > 0.0);
        let mut current = chunk + 1;
        while current <= self.chunks.len() {
            #[cfg(feature = "test-counters")]
            PREFIX_UPDATES.set(PREFIX_UPDATES.get().wrapping_add(1));
            self.chunks[current - 1].sum += delta;
            self.chunks[current - 1].count += added;
            self.chunks[current - 1].count -= removed;
            self.chunks[current - 1].sized += u32::from(is_sized && !was_sized);
            self.chunks[current - 1].sized -= u32::from(was_sized && !is_sized);
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

    /// index of the last track with a nonzero size. tracks past the last
    /// stored value take the default, so they decide unless it is zero.
    fn last_sized(&self) -> Option<u32> {
        let implicit = self.default > 0.0;
        if implicit && self.end < self.limit {
            return Some(self.limit - 1);
        }
        let sized = |node: usize| {
            let chunk = &self.chunks[node - 1];
            let unset = (node.isolate_lowest_one() * SIZE_CHUNK_LEN) as u32 - chunk.count;
            chunk.sized + if implicit { unset } else { 0 }
        };
        let mut total = 0;
        let mut index = self.chunks.len();
        while index > 0 {
            total += sized(index);
            index &= index - 1;
        }
        if total == 0 {
            return None;
        }
        let mut index = 0;
        let mut below = 0;
        let mut step = 1 << self.chunks.len().ilog2();
        while step > 0 {
            let next = index + step;
            if next <= self.chunks.len() && below + sized(next) < total {
                below += sized(next);
                index = next;
            }
            step >>= 1;
        }
        let within = self.chunks[index]
            .values
            .iter()
            .rposition(|&size| size > 0.0 || (implicit && size.is_nan()))?;
        Some((index * SIZE_CHUNK_LEN + within) as u32)
    }

    fn width(&self, index: u32) -> f32 {
        self.value(index).unwrap_or(self.default)
    }

    fn value(&self, index: u32) -> Option<f32> {
        let chunk = self.chunks.get(index as usize / SIZE_CHUNK_LEN)?;
        let value = chunk.values[index as usize % SIZE_CHUNK_LEN];
        (!value.is_nan()).then_some(value)
    }

    fn offset(&self, index: u32) -> f32 {
        let chunk = index as usize / SIZE_CHUNK_LEN;
        let (mut sum, mut count, _) = self.prefix(chunk);
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
            let (sum, count, _) = self.prefix(self.chunks.len());
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
    anchors: Arc<HashSet<(RowId, ColId)>>,
    inset: TextInset,
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
            RowFit::Screen,
        )
    }

    pub fn for_print(sheet: &Sheet, styles: &Stylesheet, metrics: &PrintMetrics) -> Self {
        Self::with_sizes(
            sheet,
            styles,
            metrics
                .default_column_width
                .map_or(metrics.base_column_pixels(), |width| {
                    metrics.column_pixels(width)
                }),
            metrics.default_row_height_pt as f64,
            NormalFace::from_metrics(metrics),
            |width| metrics.column_pixels(width),
            RowFit::Print {
                dpi: metrics.dpi as f64,
            },
        )
    }

    fn with_sizes(
        sheet: &Sheet,
        styles: &Stylesheet,
        default_col_px: f32,
        declared_row_pt: f64,
        normal: NormalFace<'_>,
        column_pixels: impl Fn(f64) -> f32,
        fit: RowFit,
    ) -> Self {
        let inset = TextInset::of(fit, normal);
        #[cfg(feature = "test-counters")]
        GEOMETRY_CONSTRUCTIONS.set(GEOMETRY_CONSTRUCTIONS.get().wrapping_add(1));
        let default_row_pt = fit.default_row(sheet, declared_row_pt, normal);
        let default_row_px = if sheet.format.zero_height {
            0.0
        } else {
            row_pt_to_px(default_row_pt)
        };
        let scale = stored_height_scale(sheet, normal, fit);
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
        let columns = Sizes::from_values(xlsx_model::MAX_COLS, default_col_px, &columns);
        let fitted = autofit_rows(
            sheet,
            styles,
            default_row_pt,
            normal,
            |col| columns.width(col),
            inset,
            fit,
        );
        let sized: std::collections::BTreeSet<_> = sheet
            .row_heights
            .keys()
            .chain(fitted.keys())
            .copied()
            .chain(
                sheet
                    .iter_cells()
                    .map(|(at, _)| at.row)
                    .chain(sheet.row_styles.keys().copied())
                    .take(if sheet.format.zero_height {
                        usize::MAX
                    } else {
                        0
                    }),
            )
            .collect();
        let rows: Vec<_> = sized
            .into_iter()
            .filter_map(|row| {
                row_size(
                    sheet,
                    row,
                    fitted.get(&row).copied(),
                    scale,
                    default_row_pt,
                    fit,
                )
                .map(|size| (row, size))
            })
            .collect();
        let rows = Sizes::from_values(xlsx_model::MAX_ROWS, default_row_px, &rows);
        Self {
            columns,
            rows,
            spanned: Arc::new(merge_anchors(sheet, true)),
            anchors: Arc::new(merge_anchors(sheet, false)),
            inset,
        }
    }

    #[doc(hidden)]
    pub fn update_merges(&mut self, sheet: &Sheet) {
        self.spanned = Arc::new(merge_anchors(sheet, true));
        self.anchors = Arc::new(merge_anchors(sheet, false));
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
        let scale = stored_height_scale(sheet, normal, RowFit::Screen);
        for row in rows {
            let mut fitted: Option<f64> = None;
            if !sheet.format.custom_height && !sheet.row_heights.contains_key(&row) {
                fitted = missing_row_style(sheet, row).and_then(|style| {
                    autofit_height(
                        styles,
                        Some(style),
                        default_pt,
                        sheet.format.default_row_height_pt.is_some(),
                        normal,
                        RowFit::Screen,
                    )
                });
                for (at, cell) in sheet.iter_cells_in_rect(row..row + 1, 0..xlsx_model::MAX_COLS) {
                    #[cfg(feature = "test-counters")]
                    AUTOFIT_CELL_VISITS.set(AUTOFIT_CELL_VISITS.get().wrapping_add(1));
                    if self.spanned.contains(&(at.row, at.col)) {
                        continue;
                    }
                    let style = sheet.effective_style(at);
                    let cached_default = sheet.format.default_row_height_pt.is_some();
                    let wrapped = wrapped_height(
                        styles,
                        style,
                        cell,
                        self.columns.width(at.col),
                        self.inset,
                        normal,
                        RowFit::Screen,
                    )
                    .filter(|height| cached_default || *height > default_pt)
                    .filter(|_| !self.anchors.contains(&(at.row, at.col)));
                    if let Some(height) = taller(
                        autofit_height(
                            styles,
                            style,
                            default_pt,
                            cached_default,
                            normal,
                            RowFit::Screen,
                        ),
                        wrapped,
                    ) {
                        fitted = Some(fitted.map_or(height, |old| old.max(height)));
                    }
                }
            }
            let size = row_size(sheet, row, fitted, scale, default_pt, RowFit::Screen);
            if self.rows.value(row) != size {
                self.rows.set(row, size);
            }
        }
    }

    /// also refits the rows whose wrapped text the resized columns rewrap.
    #[doc(hidden)]
    pub fn update_columns(
        &mut self,
        sheet: &Sheet,
        styles: &Stylesheet,
        columns: impl IntoIterator<Item = ColId>,
    ) {
        let mut resized = std::collections::BTreeSet::new();
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
                resized.insert(col);
            }
        }
        if resized.is_empty()
            || sheet.format.custom_height
            || !styles
                .cell_xfs
                .iter()
                .any(|xf| xf.alignment.as_ref().is_some_and(wrap::wraps))
        {
            return;
        }
        let rows: std::collections::BTreeSet<RowId> = sheet
            .iter_cells()
            .filter(|(at, cell)| {
                resized.contains(&at.col) && matches!(cell.value, CellValue::Text { .. })
            })
            .map(|(at, _)| at.row)
            .collect();
        self.update_rows(sheet, styles, rows);
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

    /// last row with a nonzero height; hidden trailing rows end the grid early.
    pub fn last_row_with_height(&self) -> Option<RowId> {
        self.rows.last_sized()
    }

    /// last column with a nonzero width; hidden trailing columns end the grid early.
    pub fn last_col_with_width(&self) -> Option<ColId> {
        self.columns.last_sized()
    }

    /// half-open (row, col) ranges of cells intersecting the viewport, ending at
    /// the last track with extent.
    pub fn viewport_range(&self, vp: &Viewport) -> (Range<RowId>, Range<ColId>) {
        let bottom = self.row_y(MAX_ROWS);
        let right = self.col_x(MAX_COLS);
        let rows = match self.last_row_with_height() {
            Some(last) if bottom - vp.y > 0.0 => {
                let r0 = self.row_at_y(vp.y).min(last);
                let r1 = self.row_at_y((vp.y + vp.height).min(bottom)).min(last);
                r0..r1 + 1
            }
            _ => MAX_ROWS..MAX_ROWS,
        };
        let cols = match self.last_col_with_width() {
            Some(last) if right - vp.x > 0.0 => {
                let c0 = self.col_at_x(vp.x).min(last);
                let c1 = self.col_at_x((vp.x + vp.width).min(right)).min(last);
                c0..c1 + 1
            }
            _ => MAX_COLS..MAX_COLS,
        };
        (rows, cols)
    }
}

fn row_size(
    sheet: &Sheet,
    row: RowId,
    fitted: Option<f64>,
    scale: Option<f64>,
    default_pt: f64,
    fit: RowFit,
) -> Option<f32> {
    if sheet.format.zero_height {
        return Some(if sheet.row_hidden(row) {
            0.0
        } else {
            row_pt_to_px(
                sheet
                    .row_heights
                    .get(&row)
                    .map_or(default_pt, |&height| fit.stored(height, None)),
            )
        });
    }
    sheet
        .row_heights
        .get(&row)
        .copied()
        .map(|height| fit.stored(height, scale))
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

    #[test]
    fn wrap_insets_indent_by_the_painted_indent_step() {
        let mut styles = Stylesheet::default();
        styles.fonts = vec![Font {
            name: Some("Arial".into()),
            size_pt: Some(10.0),
            ..Default::default()
        }];
        let screen = TextInset::of(RowFit::Screen, NormalFace::from_styles(&styles));
        assert_eq!(screen.indent, indent_step_px(&styles, None));
        assert_eq!(screen.indent, 12.0);
        let metrics = PrintMetrics {
            dpi: 72.0,
            max_digit_width: 5.0,
            default_row_height_pt: 12.75,
            default_column_width: None,
            font_size_pt: 10.0,
            font_family: "Arial".into(),
            font_ascent: 9.0,
            font_descent: 2.0,
        };
        let print = TextInset::of(
            RowFit::Print { dpi: 72.0 },
            NormalFace::from_metrics(&metrics),
        );
        assert_eq!(print.indent, indent_step_px(&styles, Some(&metrics)));
        assert_eq!(print.indent, 12.0);
    }

    #[test]
    fn row_default_fonts_fit_without_materializing_cells() {
        let mut sheet = Sheet::new("Rows");
        sheet.row_styles.insert(4, 1);
        let mut styles = Stylesheet::default();
        styles.fonts = vec![
            Font::default(),
            Font {
                size_pt: Some(32.0),
                ..Default::default()
            },
        ];
        styles.cell_xfs = vec![
            Xf::default(),
            Xf {
                font: Some(1),
                ..Default::default()
            },
        ];
        let mut grid = GridGeometry::new(&sheet, &styles);
        assert!(
            (grid.row_y(5) - grid.row_y(4) - row_pt_to_px(autofit_row_height_pt(32.0))).abs()
                < 0.001
        );
        assert_eq!(sheet.iter_cells().count(), 0);
        sheet.row_styles.remove(&4);
        grid.update_rows(&sheet, &styles, [4]);
        assert_eq!(grid.row_y(5), GridGeometry::new(&sheet, &styles).row_y(5));
        sheet.row_styles.insert(4, 1);
        sheet.format.zero_height = true;
        assert_eq!(
            GridGeometry::new(&sheet, &styles).row_y(5),
            row_pt_to_px(DEFAULT_ROW_HEIGHT_PT)
        );
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
                RowFit::Screen,
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
            default_row_height_pt: 15.0,
            default_column_width: None,
            font_size_pt: f32::INFINITY,
            font_family: "Calibri".into(),
            font_ascent: 14.0,
            font_descent: 4.0,
        };
        let mut sheet = Sheet::new("S");
        sheet.format.default_row_height_pt = Some(15.0);
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
            default_row_height_pt: 19.5,
            default_column_width: None,
            font_size_pt: 16.0,
            font_family: "Arial".into(),
            font_ascent: 14.0,
            font_descent: 4.0,
        };
        let mut sheet = Sheet::new("S");
        sheet.format.default_row_height_pt = Some(19.5);
        sheet.set_cell(
            CellRef::new(0, 0),
            xlsx_model::workbook::Cell {
                value: xlsx_model::value::CellValue::Number { value: 1.0 },
                ..Default::default()
            },
        );
        let grid = GridGeometry::for_print(&sheet, &Stylesheet::default(), &metrics);
        assert!((grid.row_y(1) - row_pt_to_px(20.0)).abs() < 0.001);
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

    #[test]
    fn zero_height_content_rows_use_default_height_and_hit_testing() {
        let parts = vec![
            (
                "xl/workbook.xml".into(),
                br#"<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_vec(),
            ),
            (
                "xl/_rels/workbook.xml.rels".into(),
                br#"<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>"#.to_vec(),
            ),
            (
                "xl/worksheets/sheet1.xml".into(),
                br#"<worksheet><sheetFormatPr defaultRowHeight="12.75" zeroHeight="1"/><sheetData><row r="1" ht="15"/><row r="2"><c r="A2" t="inlineStr"><is><t>visible</t></is></c></row><row r="4" hidden="1"/><row r="5"/></sheetData></worksheet>"#.to_vec(),
            ),
        ];
        let workbook = xlsx_parse::parse_workbook(&parts).unwrap();
        let sheet = &workbook.sheets[0];
        let mut grid = GridGeometry::new(sheet, &workbook.styles);
        assert_eq!(grid.row_y(1), 20.0);
        assert_eq!(grid.row_y(2) - grid.row_y(1), 17.0);
        assert_eq!(grid.row_y(3) - grid.row_y(2), 0.0);
        assert_eq!(grid.row_y(4) - grid.row_y(3), 0.0);
        assert_eq!(grid.row_y(5) - grid.row_y(4), 0.0);
        assert_eq!(grid.row_at_y(21.0), 1);
        grid.update_rows(sheet, &workbook.styles, [1, 2, 3, 4]);
        assert_eq!(grid.row_y(2), 37.0);
        assert_eq!(grid.row_y(5), 37.0);
    }

    #[test]
    fn zero_height_geometry_tracks_style_only_cells_and_row_styles() {
        let styles = Stylesheet::default();
        let mut sheet = Sheet::new("Rows");
        sheet.format.zero_height = true;
        sheet.format.default_row_height_pt = Some(12.75);
        sheet.set_cell(
            CellRef::new(0, 0),
            xlsx_model::Cell {
                style: Some(0),
                ..Default::default()
            },
        );
        sheet.row_styles.insert(1, 0);
        sheet.row_heights.insert(2, 0.0);
        sheet.set_cell(
            CellRef::new(2, 0),
            xlsx_model::Cell {
                style: Some(0),
                ..Default::default()
            },
        );
        let mut grid = GridGeometry::new(&sheet, &styles);
        assert_eq!(grid.row_y(1), 17.0);
        assert_eq!(grid.row_y(2), 34.0);
        assert_eq!(grid.row_y(3), 34.0);
        sheet.set_cell(CellRef::new(0, 0), Default::default());
        sheet.row_styles.clear();
        grid.update_rows(&sheet, &styles, [0, 1]);
        assert_eq!(grid.row_y(3), 0.0);
        sheet.row_styles.insert(3, 0);
        grid.update_rows(&sheet, &styles, [3]);
        assert_eq!(grid.row_y(4), 17.0);
        assert_eq!(grid.row_at_y(1.0), 3);
    }

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
        assert_eq!(grid.row_y(1), row_pt_to_px(DEFAULT_ROW_HEIGHT_PT));
        assert!((grid.row_y(2) - grid.row_y(1) - row_pt_to_px(24.0)).abs() < 0.001);
        assert_eq!(grid.row_y(3), grid.row_y(2));
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
            RowFit::Screen,
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
            1..4
        );
    }

    #[test]
    fn last_sized_tracks_mutations() {
        const LIMIT: u32 = 4 * SIZE_CHUNK_LEN as u32;
        let mut seed = 0x9e37_79b9_u32;
        let mut next = move || {
            seed ^= seed << 13;
            seed ^= seed >> 17;
            seed ^= seed << 5;
            seed
        };
        for default in [0.0_f32, 20.0] {
            for round in 0..48 {
                let mut values = BTreeMap::new();
                for _ in 0..next() % 12 {
                    let index = next() % LIMIT;
                    values.insert(index, if next() % 3 == 0 { 0.0 } else { 7.0 });
                }
                if round % 2 == 0 {
                    for index in next() % LIMIT..LIMIT {
                        values.insert(index, 0.0);
                    }
                }
                let pairs: Vec<_> = values.iter().map(|(&index, &size)| (index, size)).collect();
                let mut sizes = Sizes::from_values(LIMIT, default, &pairs);
                let mut model: Vec<_> = (0..LIMIT)
                    .map(|index| values.get(&index).copied())
                    .collect();
                let expected = |model: &[Option<f32>]| {
                    model
                        .iter()
                        .rposition(|size| size.unwrap_or(default) > 0.0)
                        .map(|index| index as u32)
                };
                assert_eq!(sizes.last_sized(), expected(&model), "round {round}");
                for step in 0..32 {
                    let index = if next() % 2 == 0 {
                        next() % LIMIT
                    } else {
                        (expected(&model).unwrap_or(0) + next() % 5)
                            .saturating_sub(2)
                            .min(LIMIT - 1)
                    };
                    let size = match next() % 3 {
                        0 => None,
                        1 => Some(0.0),
                        _ => Some(5.0),
                    };
                    sizes.set(index, size);
                    model[index as usize] = size;
                    assert_eq!(
                        sizes.last_sized(),
                        expected(&model),
                        "default {default} round {round} step {step}"
                    );
                }
            }
        }
    }

    #[test]
    fn viewport_range_ends_at_the_last_track_with_extent() {
        let mut sheet = Sheet::new("S");
        for col in 8..MAX_COLS {
            sheet.col_widths.insert(col, 0.0);
        }
        for row in 20..MAX_ROWS {
            sheet.row_heights.insert(row, 0.0);
        }
        let g = GridGeometry::new(&sheet, &styles());
        let right = g.col_x(MAX_COLS);
        let bottom = g.row_y(MAX_ROWS);
        let screen = Viewport {
            x: 0.0,
            y: 0.0,
            width: 1600.0,
            height: 1000.0,
        };
        assert_eq!(g.viewport_range(&screen), (0..20, 0..8));
        let corner = Viewport {
            x: right - 1.0,
            y: bottom - 1.0,
            width: 500.0,
            height: 500.0,
        };
        assert_eq!(g.viewport_range(&corner), (19..20, 7..8));
        let beyond = Viewport {
            x: right,
            y: bottom,
            ..screen
        };
        assert_eq!(
            g.viewport_range(&beyond),
            (MAX_ROWS..MAX_ROWS, MAX_COLS..MAX_COLS)
        );
    }
}
