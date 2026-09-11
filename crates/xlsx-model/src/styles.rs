//! style types (fonts, fills, borders, xf chains, theme colors). pure data;
//! the cellXfs indirection chain is walked through the `Stylesheet` accessors.

use std::collections::{HashMap, hash_map::DefaultHasher};
use std::hash::{Hash, Hasher};

use serde::{Deserialize, Serialize};

/// a color reference as it appears in a `<color>` element (§18.8.3).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Color {
    /// resolved `#rrggbb` (alpha dropped at parse; excel stores `aarrggbb`).
    Rgb(String),
    /// theme slot index (see [`Theme::slot`]) plus a signed tint in [-1.0, 1.0].
    Theme { idx: u8, tint: f64 },
    /// index into the legacy 64-entry palette.
    Indexed(u8),
    /// system/automatic color; the host decides (usually black text).
    Auto,
}

impl Color {
    /// resolve to a final `#rrggbb`, or `None` for automatic/out-of-range.
    pub fn resolve(&self, theme: &Theme) -> Option<String> {
        match self {
            Color::Rgb(s) => Some(s.clone()),
            Color::Auto => None,
            Color::Indexed(i) => indexed_color(*i).map(str::to_string),
            Color::Theme { idx, tint } => {
                let base = theme.slot(*idx)?;
                Some(apply_tint(base, *tint))
            }
        }
    }
}

/// a cell font. only the facets we render are modelled; unset fields inherit.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Font {
    pub name: Option<String>,
    pub size_pt: Option<f64>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub color: Option<Color>,
}

/// a cell fill. non-solid patterns collapse to their foreground color.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub enum Fill {
    #[default]
    None,
    Solid(Color),
}

/// border line weight/style, collapsed from the full sml set.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum BorderStyle {
    Thin,
    Medium,
    Thick,
    Dashed,
    Dotted,
    Double,
    Hair,
}

impl BorderStyle {
    /// map an sml `ST_BorderStyle` token; unknown weights fall back to `thin`.
    pub fn from_sml(s: &str) -> Self {
        match s {
            "thin" => BorderStyle::Thin,
            "medium" | "mediumDashed" | "mediumDashDot" | "mediumDashDotDot" => BorderStyle::Medium,
            "thick" => BorderStyle::Thick,
            "dashed" | "dashDot" | "dashDotDot" | "slantDashDot" => BorderStyle::Dashed,
            "dotted" => BorderStyle::Dotted,
            "double" => BorderStyle::Double,
            "hair" => BorderStyle::Hair,
            _ => BorderStyle::Thin,
        }
    }

    /// the sml token for this weight.
    pub fn as_sml(&self) -> &'static str {
        match self {
            BorderStyle::Thin => "thin",
            BorderStyle::Medium => "medium",
            BorderStyle::Thick => "thick",
            BorderStyle::Dashed => "dashed",
            BorderStyle::Dotted => "dotted",
            BorderStyle::Double => "double",
            BorderStyle::Hair => "hair",
        }
    }
}

/// one edge of a cell border.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BorderEdge {
    pub style: BorderStyle,
    pub color: Option<Color>,
}

/// the four cell edges; `None` on an edge means no border there.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Border {
    pub left: Option<BorderEdge>,
    pub right: Option<BorderEdge>,
    pub top: Option<BorderEdge>,
    pub bottom: Option<BorderEdge>,
}

/// horizontal alignment (`ST_HorizontalAlignment`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum HAlign {
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterContinuous,
    Distributed,
}

impl HAlign {
    pub fn from_sml(s: &str) -> Option<Self> {
        Some(match s {
            "general" => HAlign::General,
            "left" => HAlign::Left,
            "center" => HAlign::Center,
            "right" => HAlign::Right,
            "fill" => HAlign::Fill,
            "justify" => HAlign::Justify,
            "centerContinuous" => HAlign::CenterContinuous,
            "distributed" => HAlign::Distributed,
            _ => return None,
        })
    }

    pub fn as_sml(&self) -> &'static str {
        match self {
            HAlign::General => "general",
            HAlign::Left => "left",
            HAlign::Center => "center",
            HAlign::Right => "right",
            HAlign::Fill => "fill",
            HAlign::Justify => "justify",
            HAlign::CenterContinuous => "centerContinuous",
            HAlign::Distributed => "distributed",
        }
    }
}

/// vertical alignment (`ST_VerticalAlignment`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum VAlign {
    Top,
    Center,
    Bottom,
    Justify,
    Distributed,
}

impl VAlign {
    pub fn from_sml(s: &str) -> Option<Self> {
        Some(match s {
            "top" => VAlign::Top,
            "center" => VAlign::Center,
            "bottom" => VAlign::Bottom,
            "justify" => VAlign::Justify,
            "distributed" => VAlign::Distributed,
            _ => return None,
        })
    }

    pub fn as_sml(&self) -> &'static str {
        match self {
            VAlign::Top => "top",
            VAlign::Center => "center",
            VAlign::Bottom => "bottom",
            VAlign::Justify => "justify",
            VAlign::Distributed => "distributed",
        }
    }
}

/// cell text alignment; unset horizontal defaults to `general`, vertical to `bottom`.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default, Serialize, Deserialize)]
pub struct Alignment {
    pub h: Option<HAlign>,
    pub v: Option<VAlign>,
    pub wrap_text: bool,
    pub shrink_to_fit: bool,
}

impl Alignment {
    /// true when the element carries no meaningful alignment.
    pub fn is_empty(&self) -> bool {
        self.h.is_none() && self.v.is_none() && !self.wrap_text && !self.shrink_to_fit
    }
}

/// a cell format record (`CT_Xf` in cellXfs). a `None` pool index means
/// "inherit / default", not "index 0".
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default, Serialize, Deserialize)]
pub struct Xf {
    pub font: Option<u32>,
    pub fill: Option<u32>,
    pub border: Option<u32>,
    pub num_fmt_id: Option<u16>,
    pub alignment: Option<Alignment>,
}

/// the workbook theme colors in clrScheme declaration order
/// `[dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink]`; index via [`Theme::slot`].
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Theme {
    pub colors: [String; 12],
}

impl Default for Theme {
    fn default() -> Self {
        Theme {
            colors: [
                "#000000".into(),
                "#ffffff".into(),
                "#44546a".into(),
                "#e7e6e6".into(),
                "#4472c4".into(),
                "#ed7d31".into(),
                "#a5a5a5".into(),
                "#ffc000".into(),
                "#5b9bd5".into(),
                "#70ad47".into(),
                "#0563c1".into(),
                "#954f72".into(),
            ],
        }
    }
}

impl Theme {
    /// resolve a `theme="n"` index to a slot color. excel's index order swaps
    /// the first two light/dark pairs relative to declaration order.
    pub fn slot(&self, idx: u8) -> Option<&str> {
        let pos = match idx {
            0 => 1,
            1 => 0,
            2 => 3,
            3 => 2,
            n => n as usize,
        };
        self.colors.get(pos).map(String::as_str)
    }
}

/// hashable stand-in for a color; float tints become raw bits.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
enum ColorKey<'a> {
    Rgb(&'a str),
    Theme { idx: u8, tint: Option<u64> },
    Indexed(u8),
    Auto,
}

impl<'a> ColorKey<'a> {
    fn new(color: &'a Color) -> Option<Self> {
        Some(match color {
            Color::Rgb(s) => ColorKey::Rgb(s),
            Color::Theme { idx, tint } => ColorKey::Theme {
                idx: *idx,
                tint: Some(float_key(*tint)?),
            },
            Color::Indexed(i) => ColorKey::Indexed(*i),
            Color::Auto => ColorKey::Auto,
        })
    }
}

/// hashable stand-in for a font; `None` when a NaN size/color tint blocks keying.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct FontKey<'a> {
    name: Option<&'a str>,
    size_pt: Option<u64>,
    bold: bool,
    italic: bool,
    underline: bool,
    strike: bool,
    color: Option<ColorKey<'a>>,
}

impl<'a> FontKey<'a> {
    fn new(font: &'a Font) -> Option<Self> {
        Some(FontKey {
            name: font.name.as_deref(),
            size_pt: match font.size_pt {
                None => None,
                Some(size) => Some(float_key(size)?),
            },
            bold: font.bold,
            italic: font.italic,
            underline: font.underline,
            strike: font.strike,
            color: match &font.color {
                None => None,
                Some(color) => Some(ColorKey::new(color)?),
            },
        })
    }
}

/// hashable stand-in for a fill.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
enum FillKey<'a> {
    None,
    Solid(ColorKey<'a>),
}

impl<'a> FillKey<'a> {
    fn new(fill: &'a Fill) -> Option<Self> {
        Some(match fill {
            Fill::None => FillKey::None,
            Fill::Solid(color) => FillKey::Solid(ColorKey::new(color)?),
        })
    }
}

/// hashable stand-in for a border edge.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct EdgeKey<'a> {
    style: BorderStyle,
    color: Option<ColorKey<'a>>,
}

impl<'a> EdgeKey<'a> {
    fn new(edge: &'a BorderEdge) -> Option<Self> {
        Some(EdgeKey {
            style: edge.style,
            color: match &edge.color {
                None => None,
                Some(color) => Some(ColorKey::new(color)?),
            },
        })
    }
}

/// hashable stand-in for a border.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct BorderKey<'a> {
    left: Option<EdgeKey<'a>>,
    right: Option<EdgeKey<'a>>,
    top: Option<EdgeKey<'a>>,
    bottom: Option<EdgeKey<'a>>,
}

impl<'a> BorderKey<'a> {
    fn new(border: &'a Border) -> Option<Self> {
        Some(BorderKey {
            left: match &border.left {
                None => None,
                Some(edge) => Some(EdgeKey::new(edge)?),
            },
            right: match &border.right {
                None => None,
                Some(edge) => Some(EdgeKey::new(edge)?),
            },
            top: match &border.top {
                None => None,
                Some(edge) => Some(EdgeKey::new(edge)?),
            },
            bottom: match &border.bottom {
                None => None,
                Some(edge) => Some(EdgeKey::new(edge)?),
            },
        })
    }
}

/// intern cache: fingerprint -> index, tagged with its synced pool length.
#[derive(Debug, Default)]
struct PoolMemo {
    map: HashMap<u64, u32>,
    pool_len: usize,
    #[cfg(test)]
    rebuilds: u64,
}

/// the memo is a pure accelerator, so a clone starts cold instead of copying
/// the map; the first intern on the clone rebuilds only what it needs.
impl Clone for PoolMemo {
    fn clone(&self) -> Self {
        PoolMemo::default()
    }
}

impl PoolMemo {
    fn invalidate(&mut self, pool_len: usize) {
        self.map.clear();
        self.pool_len = pool_len;
        #[cfg(test)]
        {
            self.rebuilds += 1;
        }
    }
}

/// intern cache for `num_fmts`; hits re-verify against the live table.
#[derive(Debug, Default)]
struct FmtMemo {
    patterns: HashMap<String, (u16, usize)>,
    pool_len: usize,
    #[cfg(test)]
    rebuilds: u64,
}

impl Clone for FmtMemo {
    fn clone(&self) -> Self {
        FmtMemo::default()
    }
}

impl FmtMemo {
    fn rebuild(&mut self, formats: &[(u16, String)]) {
        self.patterns.clear();
        self.pool_len = formats.len();
        for (slot, (id, code)) in formats.iter().enumerate() {
            self.patterns.entry(code.clone()).or_insert((*id, slot));
        }
        #[cfg(test)]
        {
            self.rebuilds += 1;
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct NumFmtTableFull;

/// pool lengths captured before a batch of interning.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PoolMarks {
    fonts: usize,
    fills: usize,
    borders: usize,
    cell_xfs: usize,
    num_fmts: usize,
}

/// parsed style tables plus theme; private memos accelerate interning and
/// are invalidated by any pool length change.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Stylesheet {
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub borders: Vec<Border>,
    pub cell_xfs: Vec<Xf>,
    pub num_fmts: Vec<(u16, String)>,
    pub theme: Theme,
    #[serde(skip)]
    font_memo: PoolMemo,
    #[serde(skip)]
    fill_memo: PoolMemo,
    #[serde(skip)]
    border_memo: PoolMemo,
    #[serde(skip)]
    xf_memo: PoolMemo,
    #[serde(skip)]
    fmt_memo: FmtMemo,
}

impl PartialEq for Stylesheet {
    fn eq(&self, other: &Self) -> bool {
        self.fonts == other.fonts
            && self.fills == other.fills
            && self.borders == other.borders
            && self.cell_xfs == other.cell_xfs
            && self.num_fmts == other.num_fmts
            && self.theme == other.theme
    }
}

/// resolved number format for a cell: a custom code string or a builtin id.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FormatCode<'a> {
    Custom(&'a str),
    Builtin(u16),
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum NumberFormat {
    Builtin { id: u16 },
    Custom { pattern: String },
}

impl Default for NumberFormat {
    fn default() -> Self {
        Self::Builtin { id: 0 }
    }
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellFormat {
    pub font: Font,
    pub fill: Fill,
    pub border: Border,
    pub number_format: NumberFormat,
    pub alignment: Alignment,
}

/// borrowed view of a resolved cell format; defaults mirror [`CellFormat::default`].
#[derive(Debug, Clone, PartialEq)]
pub struct ResolvedFormat<'a> {
    pub font: &'a Font,
    pub fill: &'a Fill,
    pub border: &'a Border,
    pub number_format: FormatCode<'a>,
    pub alignment: &'a Alignment,
}

const DEFAULT_FONT: Font = Font {
    name: None,
    size_pt: None,
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    color: None,
};

const DEFAULT_BORDER: Border = Border {
    left: None,
    right: None,
    top: None,
    bottom: None,
};

const DEFAULT_ALIGNMENT: Alignment = Alignment {
    h: None,
    v: None,
    wrap_text: false,
    shrink_to_fit: false,
};

impl Stylesheet {
    /// true when no style data is present, so the serializer skips the part.
    pub fn is_empty(&self) -> bool {
        self.fonts.is_empty()
            && self.fills.is_empty()
            && self.borders.is_empty()
            && self.cell_xfs.is_empty()
            && self.num_fmts.is_empty()
            && self.theme == Theme::default()
    }

    /// the `Xf` a cell's `s` index selects.
    pub fn xf(&self, style_index: u32) -> Option<&Xf> {
        self.cell_xfs.get(style_index as usize)
    }

    /// resolved font for a cell style, or `None` when inherited/unset.
    pub fn font_for(&self, style_index: u32) -> Option<&Font> {
        let idx = self.xf(style_index)?.font?;
        self.fonts.get(idx as usize)
    }

    /// resolved fill for a cell style.
    pub fn fill_for(&self, style_index: u32) -> Option<&Fill> {
        let idx = self.xf(style_index)?.fill?;
        self.fills.get(idx as usize)
    }

    /// resolved border for a cell style.
    pub fn border_for(&self, style_index: u32) -> Option<&Border> {
        let idx = self.xf(style_index)?.border?;
        self.borders.get(idx as usize)
    }

    /// resolved alignment for a cell style.
    pub fn alignment_for(&self, style_index: u32) -> Option<&Alignment> {
        self.xf(style_index)?.alignment.as_ref()
    }

    /// the number-format for a cell style: custom code when id >= 164, else the
    /// builtin id. a missing/unset xf resolves to builtin `0` (General).
    pub fn format_code_for(&self, style_index: u32) -> FormatCode<'_> {
        let id = self
            .xf(style_index)
            .and_then(|xf| xf.num_fmt_id)
            .unwrap_or(0);
        if id >= 164
            && let Some((_, code)) = self.num_fmts.iter().find(|(k, _)| *k == id)
        {
            return FormatCode::Custom(code);
        }
        FormatCode::Builtin(id)
    }

    /// resolved format for a cell style, cloning each facet; see
    /// [`Stylesheet::resolved_format`] for the borrow-based variant.
    pub fn cell_format(&self, style_index: Option<u32>) -> CellFormat {
        let resolved = self.resolved_format(style_index);
        CellFormat {
            font: resolved.font.clone(),
            fill: resolved.fill.clone(),
            border: resolved.border.clone(),
            number_format: match resolved.number_format {
                FormatCode::Builtin(id) => NumberFormat::Builtin { id },
                FormatCode::Custom(pattern) => NumberFormat::Custom {
                    pattern: pattern.to_string(),
                },
            },
            alignment: resolved.alignment.clone(),
        }
    }

    /// borrows the resolved format for a cell style without cloning; unset
    /// facets resolve to defaults.
    pub fn resolved_format(&self, style_index: Option<u32>) -> ResolvedFormat<'_> {
        let xf = style_index.and_then(|index| self.xf(index));
        ResolvedFormat {
            font: xf
                .and_then(|xf| xf.font)
                .and_then(|index| self.fonts.get(index as usize))
                .unwrap_or(&DEFAULT_FONT),
            fill: xf
                .and_then(|xf| xf.fill)
                .and_then(|index| self.fills.get(index as usize))
                .unwrap_or(&Fill::None),
            border: xf
                .and_then(|xf| xf.border)
                .and_then(|index| self.borders.get(index as usize))
                .unwrap_or(&DEFAULT_BORDER),
            number_format: match style_index {
                Some(index) => self.format_code_for(index),
                None => FormatCode::Builtin(0),
            },
            alignment: xf
                .and_then(|xf| xf.alignment.as_ref())
                .unwrap_or(&DEFAULT_ALIGNMENT),
        }
    }

    #[inline]
    pub fn intern_cell_format(
        &mut self,
        format: &CellFormat,
    ) -> Result<Option<u32>, NumFmtTableFull> {
        if format == &CellFormat::default() {
            return Ok(None);
        }
        match &format.number_format {
            NumberFormat::Builtin { id } => Ok(Some(self.intern_builtin_cell_format(format, *id))),
            NumberFormat::Custom { pattern } => self.intern_custom_cell_format(format, pattern),
        }
    }

    #[inline(never)]
    fn intern_builtin_cell_format(&mut self, format: &CellFormat, num_fmt_id: u16) -> u32 {
        if self.cell_xfs.len() < INTERN_CACHE_MIN_POOL {
            let font = intern_linear(&mut self.fonts, &format.font);
            let fill = intern_linear(&mut self.fills, &format.fill);
            let border = intern_linear(&mut self.borders, &format.border);
            let num_fmt_id = Some(num_fmt_id).filter(|id| *id != 0);
            let alignment = (!format.alignment.is_empty()).then(|| format.alignment.clone());
            let xf = Xf {
                font: (format.font != Font::default()).then_some(font),
                fill: (format.fill != Fill::default()).then_some(fill),
                border: (format.border != Border::default()).then_some(border),
                num_fmt_id,
                alignment,
            };
            if let Some(index) = xf.font
                && xf.fill.is_none()
                && xf.border.is_none()
                && xf.num_fmt_id.is_none()
                && xf.alignment.is_none()
                && self
                    .cell_xfs
                    .get(index as usize)
                    .is_some_and(|candidate| candidate == &xf)
            {
                return index;
            }
            return intern_linear(&mut self.cell_xfs, &xf);
        }
        self.intern_cell_format_large(format, Some(num_fmt_id).filter(|id| *id != 0))
    }

    #[inline(never)]
    fn intern_custom_cell_format(
        &mut self,
        format: &CellFormat,
        pattern: &str,
    ) -> Result<Option<u32>, NumFmtTableFull> {
        let num_fmt_id = Some(self.intern_number_format(pattern).ok_or(NumFmtTableFull)?);
        if self.cell_xfs.len() < INTERN_CACHE_MIN_POOL {
            let font = intern_linear(&mut self.fonts, &format.font);
            let fill = intern_linear(&mut self.fills, &format.fill);
            let border = intern_linear(&mut self.borders, &format.border);
            let alignment = (!format.alignment.is_empty()).then(|| format.alignment.clone());
            let xf = Xf {
                font: (format.font != Font::default()).then_some(font),
                fill: (format.fill != Fill::default()).then_some(fill),
                border: (format.border != Border::default()).then_some(border),
                num_fmt_id,
                alignment,
            };
            return Ok(Some(intern_linear(&mut self.cell_xfs, &xf)));
        }
        Ok(Some(self.intern_cell_format_large(format, num_fmt_id)))
    }

    #[inline(never)]
    fn intern_cell_format_large(&mut self, format: &CellFormat, num_fmt_id: Option<u16>) -> u32 {
        if self.xf_memo.map.is_empty() {
            let font_len = self.fonts.len();
            let fill_len = self.fills.len();
            let border_len = self.borders.len();
            let font = intern_linear(&mut self.fonts, &format.font);
            let fill = intern_linear(&mut self.fills, &format.fill);
            let border = intern_linear(&mut self.borders, &format.border);
            let has_font = format.font != Font::default();
            let has_fill = format.fill != Fill::default();
            let has_border = format.border != Border::default();
            let alignment = (!format.alignment.is_empty()).then(|| format.alignment.clone());
            let xf = Xf {
                font: has_font.then_some(font),
                fill: has_fill.then_some(fill),
                border: has_border.then_some(border),
                num_fmt_id,
                alignment,
            };
            let has_new_facet = (self.fonts.len() != font_len && has_font)
                || (self.fills.len() != fill_len && has_fill)
                || (self.borders.len() != border_len && has_border);
            let index = if has_new_facet {
                self.cell_xfs.len()
            } else {
                scan_index(&self.cell_xfs, &xf)
            };
            if index == self.cell_xfs.len() {
                self.cell_xfs.push(xf);
            } else {
                seed_memo(
                    &mut self.font_memo,
                    self.fonts.len(),
                    FontKey::new(&format.font).map(|key| fingerprint(&key)),
                    font,
                );
                seed_memo(
                    &mut self.fill_memo,
                    self.fills.len(),
                    FillKey::new(&format.fill).map(|key| fingerprint(&key)),
                    fill,
                );
                seed_memo(
                    &mut self.border_memo,
                    self.borders.len(),
                    BorderKey::new(&format.border).map(|key| fingerprint(&key)),
                    border,
                );
                seed_memo(
                    &mut self.xf_memo,
                    self.cell_xfs.len(),
                    Some(fingerprint(&xf)),
                    index as u32,
                );
            }
            return index as u32;
        }
        let font = self.intern_font(&format.font);
        let fill = self.intern_fill(&format.fill);
        let border = self.intern_border(&format.border);
        let alignment = (!format.alignment.is_empty()).then(|| format.alignment.clone());
        let xf = Xf {
            font: (format.font != Font::default()).then_some(font),
            fill: (format.fill != Fill::default()).then_some(fill),
            border: (format.border != Border::default()).then_some(border),
            num_fmt_id,
            alignment,
        };
        self.intern_xf(&xf)
    }

    /// current pool lengths; interning only ever appends, so truncating back to
    /// a mark restores exactly the pools that were live when it was taken.
    pub fn pool_marks(&self) -> PoolMarks {
        PoolMarks {
            fonts: self.fonts.len(),
            fills: self.fills.len(),
            borders: self.borders.len(),
            cell_xfs: self.cell_xfs.len(),
            num_fmts: self.num_fmts.len(),
        }
    }

    /// drop every pool entry interned since `marks` was taken.
    pub fn restore_pools(&mut self, marks: PoolMarks) {
        self.fonts.truncate(marks.fonts);
        self.fills.truncate(marks.fills);
        self.borders.truncate(marks.borders);
        self.cell_xfs.truncate(marks.cell_xfs);
        self.num_fmts.truncate(marks.num_fmts);
    }

    fn intern_font(&mut self, value: &Font) -> u32 {
        intern(&mut self.fonts, &mut self.font_memo, value, |value| {
            FontKey::new(value).map(|key| fingerprint(&key))
        })
    }

    fn intern_fill(&mut self, value: &Fill) -> u32 {
        intern(&mut self.fills, &mut self.fill_memo, value, |value| {
            FillKey::new(value).map(|key| fingerprint(&key))
        })
    }

    fn intern_border(&mut self, value: &Border) -> u32 {
        intern(&mut self.borders, &mut self.border_memo, value, |value| {
            BorderKey::new(value).map(|key| fingerprint(&key))
        })
    }

    fn intern_xf(&mut self, value: &Xf) -> u32 {
        intern(&mut self.cell_xfs, &mut self.xf_memo, value, |value| {
            Some(fingerprint(value))
        })
    }

    /// true when a cell using `id` resolves to `pattern` via `format_code_for`:
    /// the id must be custom-range and its first table entry must carry the code.
    fn id_resolves_to(&self, id: u16, pattern: &str) -> bool {
        id >= 164
            && self
                .num_fmts
                .iter()
                .find(|(used, _)| *used == id)
                .is_some_and(|(_, code)| code == pattern)
    }

    fn intern_number_format(&mut self, pattern: &str) -> Option<u16> {
        if self.fmt_memo.pool_len != self.num_fmts.len() {
            self.fmt_memo.rebuild(&self.num_fmts);
        }
        if let Some(&(id, slot)) = self.fmt_memo.patterns.get(pattern) {
            let slot_matches = self
                .num_fmts
                .get(slot)
                .is_some_and(|(used, code)| *used == id && code == pattern);
            let resolves = self.id_resolves_to(id, pattern);
            if slot_matches && resolves {
                return Some(id);
            }
            self.fmt_memo.rebuild(&self.num_fmts);
            if let Some(&(id, _)) = self.fmt_memo.patterns.get(pattern)
                && self.id_resolves_to(id, pattern)
            {
                return Some(id);
            }
        }
        if let Some((slot, (id, _))) = self
            .num_fmts
            .iter()
            .enumerate()
            .find(|(_, (id, code))| code == pattern && self.id_resolves_to(*id, pattern))
        {
            self.fmt_memo
                .patterns
                .insert(pattern.to_string(), (*id, slot));
            return Some(*id);
        }
        let live_ids: std::collections::HashSet<u16> =
            self.num_fmts.iter().map(|(id, _)| *id).collect();
        let candidate = (164..=u16::MAX).find(|id| !live_ids.contains(id))?;
        let slot = self.num_fmts.len();
        self.num_fmts.push((candidate, pattern.to_string()));
        self.fmt_memo
            .patterns
            .insert(pattern.to_string(), (candidate, slot));
        self.fmt_memo.pool_len = self.num_fmts.len();
        Some(candidate)
    }
}

const INTERN_CACHE_MIN_POOL: usize = 128;

fn seed_memo(memo: &mut PoolMemo, pool_len: usize, key: Option<u64>, index: u32) {
    if let Some(key) = key {
        memo.map.insert(key, index);
    }
    memo.pool_len = pool_len;
}

fn intern<T>(
    values: &mut Vec<T>,
    memo: &mut PoolMemo,
    value: &T,
    make_key: impl FnOnce(&T) -> Option<u64>,
) -> u32
where
    T: PartialEq + Clone,
{
    if values.len() < INTERN_CACHE_MIN_POOL {
        return intern_linear(values, value);
    }

    if !memo.map.is_empty() && memo.pool_len != values.len() {
        memo.invalidate(values.len());
    }

    if memo.map.is_empty() {
        let index = scan_index(values, value);
        if index == values.len() {
            values.push(value.clone());
        } else if values.len() >= INTERN_CACHE_MIN_POOL
            && let Some(key) = make_key(value)
        {
            memo.map.insert(key, index as u32);
            memo.pool_len = values.len();
        }
        return index as u32;
    }

    let Some(key) = make_key(value) else {
        let index = scan_index(values, value);
        if index == values.len() {
            values.push(value.clone());
            memo.pool_len = values.len();
        }
        return index as u32;
    };
    let cached = memo.map.get(&key).copied();
    if let Some(index) = cached
        && values
            .get(index as usize)
            .is_some_and(|candidate| candidate == value)
    {
        return index;
    }
    if let Some(found) = values.iter().position(|candidate| candidate == value) {
        memo.map.insert(key, found as u32);
        return found as u32;
    }
    let index = values.len() as u32;
    values.push(value.clone());
    if cached.is_some() {
        memo.map.insert(key, index);
    }
    memo.pool_len = values.len();
    index
}

#[inline(always)]
fn intern_linear<T: PartialEq + Clone>(values: &mut Vec<T>, value: &T) -> u32 {
    for (index, candidate) in values.iter().enumerate() {
        if candidate == value {
            return index as u32;
        }
    }
    values.push(value.clone());
    (values.len() - 1) as u32
}

fn scan_index<T: PartialEq>(values: &[T], value: &T) -> usize {
    for (index, candidate) in values.iter().enumerate() {
        if candidate == value {
            return index;
        }
    }
    values.len()
}

fn fingerprint(value: &impl Hash) -> u64 {
    let mut hasher = DefaultHasher::new();
    value.hash(&mut hasher);
    hasher.finish()
}

/// hashable stand-in for an f64; `None` for NaN (which never compares equal)
/// with -0.0 folded onto 0.0 so bits equality matches float equality.
fn float_key(v: f64) -> Option<u64> {
    if v.is_nan() {
        None
    } else {
        Some(if v == 0.0 { 0.0 } else { v }.to_bits())
    }
}

/// apply a spreadsheetml tint to a `#rrggbb` in hsl luminance per §18.8.3:
/// negative darkens (`L*(1+tint)`), positive lightens (`L*(1-tint)+tint`).
fn apply_tint(hex: &str, tint: f64) -> String {
    let (r, g, b) = match parse_hex(hex) {
        Some(rgb) => rgb,
        None => return hex.to_string(),
    };
    if tint == 0.0 {
        return format!("#{r:02x}{g:02x}{b:02x}");
    }
    let (h, s, l) = rgb_to_hsl(r, g, b);
    let l2 = if tint < 0.0 {
        l * (1.0 + tint)
    } else {
        l * (1.0 - tint) + tint
    }
    .clamp(0.0, 1.0);
    let (r2, g2, b2) = hsl_to_rgb(h, s, l2);
    format!("#{r2:02x}{g2:02x}{b2:02x}")
}

/// parse `#rrggbb` (leading `#` optional) into `(r, g, b)` bytes.
fn parse_hex(hex: &str) -> Option<(u8, u8, u8)> {
    let h = hex.strip_prefix('#').unwrap_or(hex);
    if h.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&h[0..2], 16).ok()?;
    let g = u8::from_str_radix(&h[2..4], 16).ok()?;
    let b = u8::from_str_radix(&h[4..6], 16).ok()?;
    Some((r, g, b))
}

/// rgb bytes -> hsl with hue in degrees [0,360) and s/l in [0,1].
fn rgb_to_hsl(r: u8, g: u8, b: u8) -> (f64, f64, f64) {
    let (r, g, b) = (r as f64 / 255.0, g as f64 / 255.0, b as f64 / 255.0);
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    let d = max - min;
    if d == 0.0 {
        return (0.0, 0.0, l);
    }
    let s = d / (1.0 - (2.0 * l - 1.0).abs());
    let h = if max == r {
        60.0 * (((g - b) / d).rem_euclid(6.0))
    } else if max == g {
        60.0 * ((b - r) / d + 2.0)
    } else {
        60.0 * ((r - g) / d + 4.0)
    };
    (h, s, l)
}

/// hsl (hue degrees, s/l in [0,1]) -> rgb bytes, rounding each channel.
fn hsl_to_rgb(h: f64, s: f64, l: f64) -> (u8, u8, u8) {
    let c = (1.0 - (2.0 * l - 1.0).abs()) * s;
    let hp = h / 60.0;
    let x = c * (1.0 - (hp.rem_euclid(2.0) - 1.0).abs());
    let (r1, g1, b1) = match hp as i32 {
        0 => (c, x, 0.0),
        1 => (x, c, 0.0),
        2 => (0.0, c, x),
        3 => (0.0, x, c),
        4 => (x, 0.0, c),
        _ => (c, 0.0, x),
    };
    let m = l - c / 2.0;
    let to_byte = |v: f64| ((v + m) * 255.0).round().clamp(0.0, 255.0) as u8;
    (to_byte(r1), to_byte(g1), to_byte(b1))
}

/// the legacy 64-entry indexed palette (biff8 default); system indices 64/65
/// resolve to `None`.
fn indexed_color(i: u8) -> Option<&'static str> {
    const PALETTE: [&str; 64] = [
        "#000000", "#ffffff", "#ff0000", "#00ff00", "#0000ff", "#ffff00", "#ff00ff", "#00ffff",
        "#000000", "#ffffff", "#ff0000", "#00ff00", "#0000ff", "#ffff00", "#ff00ff", "#00ffff",
        "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#c0c0c0", "#808080",
        "#9999ff", "#993366", "#ffffcc", "#ccffff", "#660066", "#ff8080", "#0066cc", "#ccccff",
        "#000080", "#ff00ff", "#ffff00", "#00ffff", "#800080", "#800000", "#008080", "#0000ff",
        "#00ccff", "#ccffff", "#ccffcc", "#ffff99", "#99ccff", "#ff99cc", "#cc99ff", "#ffcc99",
        "#3366ff", "#33cccc", "#99cc00", "#ffcc00", "#ff9900", "#ff6600", "#666699", "#969696",
        "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
    ];
    PALETTE.get(i as usize).copied()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn theme_slot_swaps_first_two_pairs() {
        let t = Theme::default();
        assert_eq!(t.slot(0), Some("#ffffff"));
        assert_eq!(t.slot(1), Some("#000000"));
        assert_eq!(t.slot(2), Some("#e7e6e6"));
        assert_eq!(t.slot(3), Some("#44546a"));
        assert_eq!(t.slot(4), Some("#4472c4"));
        assert_eq!(t.slot(11), Some("#954f72"));
        assert_eq!(t.slot(12), None);
    }

    #[test]
    fn tint_zero_is_identity() {
        let t = Theme::default();
        let c = Color::Theme { idx: 4, tint: 0.0 };
        assert_eq!(c.resolve(&t).as_deref(), Some("#4472c4"));
    }

    #[test]
    fn tint_negative_darkens_accent1_matches_excel() {
        let t = Theme::default();
        let c = Color::Theme {
            idx: 4,
            tint: -0.25,
        };
        assert_eq!(c.resolve(&t).as_deref(), Some("#2f5597"));
    }

    #[test]
    fn tint_positive_lightens_accent1_matches_excel() {
        let t = Theme::default();
        let c = Color::Theme { idx: 4, tint: 0.4 };
        assert_eq!(c.resolve(&t).as_deref(), Some("#8faadc"));
    }

    #[test]
    fn indexed_and_rgb_and_auto_resolve() {
        let t = Theme::default();
        assert_eq!(Color::Indexed(2).resolve(&t).as_deref(), Some("#ff0000"));
        assert_eq!(Color::Indexed(64).resolve(&t), None);
        assert_eq!(
            Color::Rgb("#123456".into()).resolve(&t).as_deref(),
            Some("#123456")
        );
        assert_eq!(Color::Auto.resolve(&t), None);
    }

    #[test]
    fn accessors_walk_the_indirection_chain() {
        let mut ss = Stylesheet {
            fonts: vec![
                Font::default(),
                Font {
                    bold: true,
                    ..Font::default()
                },
            ],
            fills: vec![Fill::None, Fill::Solid(Color::Rgb("#ffff00".into()))],
            borders: vec![Border::default()],
            cell_xfs: vec![
                Xf::default(),
                Xf {
                    font: Some(1),
                    fill: Some(1),
                    border: Some(0),
                    num_fmt_id: Some(164),
                    alignment: Some(Alignment {
                        h: Some(HAlign::Center),
                        v: Some(VAlign::Center),
                        wrap_text: true,
                        shrink_to_fit: false,
                    }),
                },
            ],
            num_fmts: vec![(164, "0.0\"%\"".into())],
            theme: Theme::default(),
            ..Stylesheet::default()
        };

        assert!(ss.font_for(1).unwrap().bold);
        assert_eq!(
            ss.fill_for(1),
            Some(&Fill::Solid(Color::Rgb("#ffff00".into())))
        );
        assert_eq!(ss.border_for(1), Some(&Border::default()));
        assert_eq!(ss.format_code_for(1), FormatCode::Custom("0.0\"%\""));
        assert_eq!(ss.format_code_for(0), FormatCode::Builtin(0));
        assert!(ss.font_for(0).is_none());
        assert!(ss.xf(99).is_none());
        assert_eq!(ss.format_code_for(99), FormatCode::Builtin(0));

        ss.cell_xfs[1].num_fmt_id = Some(14);
        assert_eq!(ss.format_code_for(1), FormatCode::Builtin(14));
    }

    #[test]
    fn cell_formats_intern_and_resolve() {
        let mut styles = Stylesheet::default();
        let format = CellFormat {
            font: Font {
                name: Some("Arial".into()),
                bold: true,
                color: Some(Color::Rgb("#123456".into())),
                ..Font::default()
            },
            number_format: NumberFormat::Custom {
                pattern: "0.000".into(),
            },
            alignment: Alignment {
                h: Some(HAlign::Center),
                wrap_text: true,
                ..Alignment::default()
            },
            ..CellFormat::default()
        };
        let first = styles.intern_cell_format(&format).unwrap().unwrap();
        let second = styles.intern_cell_format(&format).unwrap().unwrap();
        assert_eq!(first, second);
        assert_eq!(styles.cell_format(Some(first)), format);
        assert_eq!(styles.intern_cell_format(&CellFormat::default()), Ok(None));
    }

    #[test]
    fn pool_interning_is_stable_and_reuses_indices() {
        let mut ss = Stylesheet::default();
        let plain = Font {
            name: Some("Calibri".into()),
            size_pt: Some(11.0),
            ..Font::default()
        };
        let bold = Font {
            bold: true,
            ..Font::default()
        };
        assert_eq!(ss.intern_font(&plain), 0);
        assert_eq!(ss.intern_font(&bold), 1);
        assert_eq!(ss.intern_font(&plain), 0);
        assert_eq!(ss.fonts.len(), 2);

        let yellow = Fill::Solid(Color::Rgb("#ffff00".into()));
        assert_eq!(ss.intern_fill(&Fill::None), 0);
        assert_eq!(ss.intern_fill(&yellow), 1);
        assert_eq!(ss.intern_fill(&yellow), 1);

        let boxed = Border {
            top: Some(BorderEdge {
                style: BorderStyle::Thin,
                color: None,
            }),
            ..Border::default()
        };
        let dashed = Border {
            bottom: Some(BorderEdge {
                style: BorderStyle::Dashed,
                color: Some(Color::Indexed(3)),
            }),
            ..Border::default()
        };
        assert_eq!(ss.intern_border(&Border::default()), 0);
        assert_eq!(ss.intern_border(&boxed), 1);
        assert_eq!(ss.intern_border(&dashed), 2);
        assert_eq!(ss.intern_border(&boxed), 1);
        assert_eq!(ss.borders.len(), 3);

        let xf_a = Xf {
            font: Some(0),
            fill: Some(1),
            ..Xf::default()
        };
        let xf_b = Xf {
            border: Some(2),
            alignment: Some(Alignment {
                h: Some(HAlign::Right),
                ..Alignment::default()
            }),
            ..Xf::default()
        };
        assert_eq!(ss.intern_xf(&xf_a), 0);
        assert_eq!(ss.intern_xf(&xf_b), 1);
        assert_eq!(ss.intern_xf(&xf_a), 0);
        assert_eq!(ss.cell_xfs.len(), 2);
    }

    #[test]
    fn font_key_covers_every_font_equality_field() {
        let base = Font::default();
        let variants = [
            Font {
                name: Some("Aptos".into()),
                ..base.clone()
            },
            Font {
                size_pt: Some(11.0),
                ..base.clone()
            },
            Font {
                bold: true,
                ..base.clone()
            },
            Font {
                italic: true,
                ..base.clone()
            },
            Font {
                underline: true,
                ..base.clone()
            },
            Font {
                strike: true,
                ..base.clone()
            },
            Font {
                color: Some(Color::Rgb("#123456".into())),
                ..base.clone()
            },
        ];
        let base_key = FontKey::new(&base).unwrap();
        for variant in variants {
            assert_ne!(FontKey::new(&variant).unwrap(), base_key);
        }
    }

    #[test]
    fn adaptive_interning_skips_small_and_unique_pool_keys() {
        let mut small = Stylesheet::default();
        let common = Font {
            name: Some("Aptos".into()),
            ..Font::default()
        };
        for _ in 0..100 {
            assert_eq!(small.intern_font(&common), 0);
        }
        assert!(small.font_memo.map.is_empty());

        let mut unique = Stylesheet::default();
        let mut formats = Vec::new();
        for index in 0..(INTERN_CACHE_MIN_POOL + 8) {
            let format = CellFormat {
                font: Font {
                    name: Some(format!("font-{index}")),
                    ..Font::default()
                },
                ..CellFormat::default()
            };
            assert_eq!(
                unique.intern_cell_format(&format).unwrap(),
                Some(index as u32)
            );
            formats.push(format);
        }
        assert!(unique.font_memo.map.is_empty());
        assert!(unique.xf_memo.map.is_empty());

        let repeated = formats.last().unwrap();
        let repeated_index = (formats.len() - 1) as u32;
        assert_eq!(
            unique.intern_cell_format(repeated).unwrap(),
            Some(repeated_index)
        );
        assert_eq!(unique.font_memo.map.len(), 1);
        assert_eq!(unique.xf_memo.map.len(), 1);
        assert_eq!(
            unique.intern_cell_format(repeated).unwrap(),
            Some(repeated_index)
        );
        assert_eq!(unique.font_memo.map.len(), 1);
        assert_eq!(unique.xf_memo.map.len(), 1);
    }

    #[test]
    fn number_format_interning_reuses_ids_and_allocates_lowest_free() {
        let mut ss = Stylesheet::default();
        assert_eq!(ss.intern_number_format("0.000"), Some(164));
        assert_eq!(ss.intern_number_format("#,##0"), Some(165));
        assert_eq!(ss.intern_number_format("0.000"), Some(164));
        assert_eq!(ss.num_fmts.len(), 2);

        ss.num_fmts.clear();
        ss.num_fmts.push((165, "#,##0".into()));
        assert_eq!(ss.intern_number_format("0.000"), Some(164));
        assert_eq!(ss.intern_number_format("#,##0"), Some(165));
    }

    #[test]
    fn in_place_pool_edit_self_heals_on_next_intern() {
        let mut ss = Stylesheet::default();
        let a = Font {
            name: Some("A".into()),
            ..Font::default()
        };
        let b = Font {
            name: Some("B".into()),
            ..Font::default()
        };
        assert_eq!(ss.intern_font(&a), 0);
        ss.fonts[0] = b.clone();

        assert_eq!(ss.intern_font(&a), 1);
        assert_eq!(ss.fonts.len(), 2);
        assert_eq!(ss.fonts[1], a);

        assert_eq!(ss.intern_font(&b), 0);
        assert_eq!(ss.fonts.len(), 2);
    }

    #[test]
    fn same_length_swap_dedups_via_scan_instead_of_appending() {
        let mut ss = Stylesheet::default();
        let a = Font {
            name: Some("A".into()),
            ..Font::default()
        };
        let b = Font {
            name: Some("B".into()),
            ..Font::default()
        };
        assert_eq!(ss.intern_font(&a), 0);
        ss.fonts[0] = b.clone();

        assert_eq!(
            ss.intern_font(&b),
            0,
            "memo miss must rescan before appending"
        );
        assert_eq!(ss.fonts.len(), 1);

        assert_eq!(ss.intern_font(&a), 1);
        assert_eq!(ss.fonts, vec![b, a]);
    }

    #[test]
    fn duplicate_and_nan_entries_do_not_force_rebuilds() {
        let mut ss = Stylesheet::default();
        let plain = Font {
            name: Some("A".into()),
            ..Font::default()
        };
        let nan = Font {
            size_pt: Some(f64::NAN),
            ..Font::default()
        };
        ss.fonts = vec![plain.clone(), plain.clone(), nan.clone()];

        assert_eq!(ss.intern_font(&plain), 0);
        assert_eq!(ss.font_memo.rebuilds, 0);

        assert_eq!(ss.intern_font(&plain), 0);
        assert_eq!(ss.font_memo.rebuilds, 0);

        let first = ss.intern_font(&nan);
        let second = ss.intern_font(&nan);
        assert_ne!(first, second);
        assert_eq!(first, 3);
        assert_eq!(second, 4);

        let stable = ss.font_memo.rebuilds;
        assert_eq!(ss.intern_font(&Font::default()), 5);
        assert_eq!(ss.font_memo.rebuilds, stable);
        assert_eq!(ss.intern_font(&Font::default()), 5);
        assert_eq!(ss.font_memo.rebuilds, stable);
        assert_eq!(ss.fonts.len(), 6);
    }

    #[test]
    fn shrink_then_intern_does_not_panic_and_returns_fresh_indices() {
        let mut ss = Stylesheet::default();
        let a = Font {
            name: Some("A".into()),
            ..Font::default()
        };
        let b = Font {
            name: Some("B".into()),
            ..Font::default()
        };
        let nan = Font {
            size_pt: Some(f64::NAN),
            ..Font::default()
        };
        assert_eq!(ss.intern_font(&a), 0);
        assert_eq!(ss.intern_font(&b), 1);

        ss.fonts.clear();
        assert_eq!(ss.intern_font(&nan), 0);
        assert_eq!(ss.intern_font(&b), 1);
        assert_eq!(ss.intern_font(&a), 2);
        assert_eq!(ss.fonts.len(), 3);
        assert!(ss.fonts[0].size_pt.unwrap().is_nan());
        assert_eq!(ss.fonts[1], b);
        assert_eq!(ss.fonts[2], a);
    }

    #[test]
    fn same_length_num_fmt_swap_rebuilds_and_avoids_collisions() {
        let mut ss = Stylesheet::default();
        assert_eq!(ss.intern_number_format("a"), Some(164));
        assert_eq!(ss.intern_number_format("b"), Some(165));

        ss.num_fmts = vec![(164, "x".into()), (170, "y".into())];
        assert_eq!(ss.intern_number_format("a"), Some(165));
        assert_eq!(
            ss.num_fmts,
            [(164, "x".into()), (170, "y".into()), (165, "a".into())]
        );
        assert_eq!(ss.intern_number_format("y"), Some(170));
        assert_eq!(ss.intern_number_format("x"), Some(164));

        ss.num_fmts = vec![(200, "k".into())];
        assert_eq!(ss.intern_number_format("q"), Some(164));
        assert_eq!(ss.num_fmts, [(200, "k".into()), (164, "q".into())]);

        ss.num_fmts = vec![(164, "z".into())];
        assert_eq!(ss.intern_number_format("q"), Some(165));
        assert_eq!(ss.num_fmts, [(164, "z".into()), (165, "q".into())]);

        ss.num_fmts = vec![(166, "m".into()), (167, "n".into())];
        assert_eq!(ss.intern_number_format("qq"), Some(164));
        assert_eq!(
            ss.num_fmts,
            [(166, "m".into()), (167, "n".into()), (164, "qq".into())]
        );
    }

    #[test]
    fn number_format_allocation_fails_explicitly_when_ids_are_exhausted() {
        let mut ss = Stylesheet {
            num_fmts: (164..=u16::MAX).map(|id| (id, format!("p{id}"))).collect(),
            ..Stylesheet::default()
        };
        assert_eq!(ss.intern_number_format("overflow"), None);
        assert!(ss.num_fmts.len() == (u16::MAX - 164 + 1) as usize);
    }

    #[test]
    fn fmt_memo_miss_scans_live_table_before_allocating() {
        let mut ss = Stylesheet::default();
        assert_eq!(ss.intern_number_format("0.0"), Some(164));
        ss.fmt_memo.patterns.clear();
        assert_eq!(ss.intern_number_format("0.0"), Some(164));
        assert_eq!(ss.num_fmts.len(), 1);
    }

    #[test]
    fn absent_pattern_allocation_validates_candidate_against_live_table() {
        let mut ss = Stylesheet {
            num_fmts: vec![(164, "a".into()), (200, "b".into())],
            ..Stylesheet::default()
        };
        assert_eq!(ss.intern_number_format("m"), Some(165));

        ss.num_fmts[1] = (166, "evil".into());
        let allocated = ss.intern_number_format("fresh");
        let live_ids: Vec<u16> = ss.num_fmts.iter().map(|(id, _)| *id).collect();
        let Some(allocated) = allocated else {
            panic!("allocation must succeed while ids remain");
        };
        assert_ne!(
            allocated, 166,
            "allocated id must not collide with a live id"
        );
        assert_eq!(allocated, 167);
        assert!(!live_ids[..3].contains(&allocated));
        assert_eq!(
            ss.num_fmts,
            [
                (164, "a".into()),
                (166, "evil".into()),
                (165, "m".into()),
                (167, "fresh".into()),
            ]
        );

        ss.num_fmts = vec![(164, "General".into())];
        assert_eq!(ss.intern_number_format("m"), Some(165));
        ss.num_fmts[1] = (164, "m".into());
        assert_eq!(ss.intern_number_format("q"), Some(165));
        assert_eq!(
            ss.num_fmts,
            [
                (164, "General".into()),
                (164, "m".into()),
                (165, "q".into()),
            ]
        );
    }

    #[test]
    fn duplicate_num_fmt_ids_intern_to_round_trippable_ids() {
        let mut ss = Stylesheet {
            num_fmts: vec![
                (164, "a".into()),
                (164, "b".into()),
                (100, "c".into()),
                (300, "d".into()),
                (300, "e".into()),
            ],
            ..Stylesheet::default()
        };
        let mut intern_and_check = |pattern: &str| {
            let id = ss.intern_number_format(pattern).unwrap();
            let idx = ss.cell_xfs.len() as u32;
            ss.cell_xfs.push(Xf {
                num_fmt_id: Some(id),
                ..Xf::default()
            });
            assert_eq!(
                ss.format_code_for(idx),
                FormatCode::Custom(pattern),
                "interned id must resolve back to the pattern"
            );
            id
        };

        assert_eq!(intern_and_check("a"), 164);
        assert_eq!(
            intern_and_check("b"),
            165,
            "duplicate id must not be reused"
        );
        assert_eq!(intern_and_check("c"), 166, "sub-spec id must not be reused");
        assert_eq!(intern_and_check("d"), 300);
        assert_eq!(
            intern_and_check("e"),
            167,
            "duplicate id must not be reused"
        );

        assert_eq!(intern_and_check("a"), 164);
        assert_eq!(intern_and_check("b"), 165);
        assert_eq!(intern_and_check("c"), 166);
        assert_eq!(intern_and_check("d"), 300);
        assert_eq!(intern_and_check("e"), 167);
        assert_eq!(
            ss.num_fmts,
            [
                (164, "a".into()),
                (164, "b".into()),
                (100, "c".into()),
                (300, "d".into()),
                (300, "e".into()),
                (165, "b".into()),
                (166, "c".into()),
                (167, "e".into()),
            ]
        );
    }

    #[test]
    fn resolved_format_borrows_without_defaulting_gaps() {
        let mut ss = Stylesheet::default();
        let format = CellFormat {
            font: Font {
                bold: true,
                size_pt: Some(-0.0),
                ..Font::default()
            },
            number_format: NumberFormat::Custom {
                pattern: "0.0".into(),
            },
            ..CellFormat::default()
        };
        let index = ss.intern_cell_format(&format).unwrap().unwrap();

        let resolved = ss.resolved_format(Some(index));
        assert!(resolved.font.bold);
        assert_eq!(resolved.font.size_pt, Some(-0.0));
        assert_eq!(
            resolved.number_format,
            FormatCode::Custom("0.0"),
            "custom code borrows the pooled string"
        );
        assert_eq!(resolved.fill, &Fill::None);
        assert!(!resolved.alignment.wrap_text);

        let fallback = ss.resolved_format(Some(99));
        assert_eq!(fallback.font, &Font::default());
        assert_eq!(fallback.number_format, FormatCode::Builtin(0));
        assert_eq!(
            ss.resolved_format(None).number_format,
            FormatCode::Builtin(0)
        );
        assert_eq!(ss.cell_format(Some(index)), format);
    }

    #[test]
    fn stylesheet_round_trips_and_keeps_indices_after_reload() {
        let mut styles = Stylesheet::default();
        let warm = CellFormat {
            font: Font {
                italic: true,
                color: Some(Color::Theme {
                    idx: 4,
                    tint: -0.25,
                }),
                ..Font::default()
            },
            number_format: NumberFormat::Custom {
                pattern: "$#,##0.00".into(),
            },
            alignment: Alignment {
                v: Some(VAlign::Top),
                shrink_to_fit: true,
                ..Alignment::default()
            },
            ..CellFormat::default()
        };
        let cold = CellFormat {
            border: Border {
                left: Some(BorderEdge {
                    style: BorderStyle::Double,
                    color: Some(Color::Rgb("#123456".into())),
                }),
                ..Border::default()
            },
            ..CellFormat::default()
        };
        let warm_index = styles.intern_cell_format(&warm).unwrap().unwrap();
        let cold_index = styles.intern_cell_format(&cold).unwrap().unwrap();
        assert_ne!(warm_index, cold_index);

        let json = serde_json::to_string(&styles).unwrap();
        let mut reloaded: Stylesheet = serde_json::from_str(&json).unwrap();
        assert_eq!(reloaded, styles);
        assert_eq!(serde_json::to_string(&reloaded).unwrap(), json);

        assert_eq!(
            reloaded.intern_cell_format(&warm).unwrap().unwrap(),
            warm_index
        );
        assert_eq!(
            reloaded.intern_cell_format(&cold).unwrap().unwrap(),
            cold_index
        );
        assert_eq!(reloaded.cell_format(Some(warm_index)), warm);
        assert_eq!(reloaded.num_fmts.len(), 1);
    }

    #[test]
    fn a_clone_starts_cold_and_still_dedups_against_the_pools() {
        let mut styles = Stylesheet::default();
        let format = CellFormat {
            font: Font {
                bold: true,
                name: Some("Calibri".into()),
                ..Font::default()
            },
            fill: Fill::Solid(Color::Rgb("#ff0000".into())),
            number_format: NumberFormat::Custom {
                pattern: "0.00".into(),
            },
            ..CellFormat::default()
        };
        let index = styles.intern_cell_format(&format).unwrap().unwrap();

        let mut cloned = styles.clone();
        assert!(cloned.font_memo.map.is_empty(), "clone must start cold");
        assert!(cloned.fmt_memo.patterns.is_empty(), "clone must start cold");
        assert_eq!(cloned, styles, "a cold memo must not change equality");

        assert_eq!(cloned.intern_cell_format(&format).unwrap(), Some(index));
        assert_eq!(cloned.fonts.len(), styles.fonts.len());
        assert_eq!(cloned.fills.len(), styles.fills.len());
        assert_eq!(cloned.cell_xfs.len(), styles.cell_xfs.len());
        assert_eq!(cloned.num_fmts.len(), styles.num_fmts.len());
        assert_eq!(cloned, styles);
    }

    #[test]
    fn restore_pools_drops_everything_interned_since_the_mark() {
        let mut ss = Stylesheet::default();
        let first = CellFormat {
            font: Font {
                bold: true,
                ..Font::default()
            },
            number_format: NumberFormat::Custom {
                pattern: "0.0".into(),
            },
            ..CellFormat::default()
        };
        ss.intern_cell_format(&first).unwrap().unwrap();
        let marks = ss.pool_marks();
        let before = ss.clone();

        let second = CellFormat {
            font: Font {
                italic: true,
                name: Some("Arial".into()),
                ..Font::default()
            },
            fill: Fill::Solid(Color::Indexed(7)),
            number_format: NumberFormat::Custom {
                pattern: "#,##0".into(),
            },
            ..CellFormat::default()
        };
        ss.intern_cell_format(&second).unwrap().unwrap();
        assert_ne!(ss, before);

        ss.restore_pools(marks);
        assert_eq!(ss, before, "truncating to a mark must restore the pools");
        assert_eq!(
            ss.intern_cell_format(&first).unwrap(),
            before.cell_xfs.len().checked_sub(1).map(|i| i as u32),
            "the memo must not hand back a truncated index"
        );
        assert_eq!(ss, before);
    }

    #[test]
    fn nan_sizes_fall_back_without_deduping() {
        let mut ss = Stylesheet::default();
        let nan = Font {
            size_pt: Some(f64::NAN),
            ..Font::default()
        };
        let a = ss.intern_font(&nan);
        let b = ss.intern_font(&nan);
        assert_eq!(a, 0);
        assert_ne!(a, b);
        assert_eq!(ss.fonts.len(), 2);
        assert_eq!(ss.intern_font(&Font::default()), 2);
        let with_nan_tint = Font {
            color: Some(Color::Theme {
                idx: 4,
                tint: f64::NAN,
            }),
            ..Font::default()
        };
        assert_eq!(ss.intern_font(&with_nan_tint), 3);
        assert_eq!(ss.intern_font(&nan), 4);
    }
}
